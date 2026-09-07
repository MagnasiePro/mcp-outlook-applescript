import { z } from 'zod';
import type { IMailSender, MailSenderParams } from '../applescript/mail-sender.js';
import { ApprovalTokenManager, hashSendPayloadForApproval } from '../approval/index.js';
import type { ValidationErrorReason } from '../approval/index.js';
import { assertSafeAttachmentPath } from '../utils/paths.js';
import { logAudit } from '../utils/audit.js';
import { SendEmailInput, type SendEmailParams } from './mail.js';
import {
    ApprovalExpiredError,
    ApprovalInvalidError,
    TargetChangedError,
} from '../utils/errors.js';

// ---------------------------------------------------------------------------
// Input schemas — two-phase send
// ---------------------------------------------------------------------------

/** Phase 1: same payload as the old one-shot send_email. Returns an approval token. */
export const PrepareSendEmailInput = SendEmailInput;

/** Phase 2: execute a previously prepared send using its approval token. */
export const ConfirmSendEmailInput = z.strictObject({
    token_id: z.uuid().describe('The approval token UUID returned by prepare_send_email (e.g., "a1b2c3d4-...")'),
});

/** Validated parameters for confirming a send. */
export type ConfirmSendEmailParams = z.infer<typeof ConfirmSendEmailInput>;

/** Sends have no pre-existing entity, so the token is bound to a synthetic target id. */
const SEND_TARGET_ID = 0;
/** Max characters of body echoed back in the prepare preview. */
const BODY_PREVIEW_LIMIT = 500;

/** Maps an approval validation failure reason to a typed error. */
function throwValidationError(reason: ValidationErrorReason | undefined): never {
    switch (reason) {
        case 'EXPIRED':
            throw new ApprovalExpiredError();
        case 'NOT_FOUND':
            throw new ApprovalInvalidError('Token not found or already used');
        case 'OPERATION_MISMATCH':
            throw new ApprovalInvalidError('Token was issued for a different operation');
        case 'TARGET_MISMATCH':
            throw new ApprovalInvalidError('Token was issued for a different target');
        case 'ALREADY_CONSUMED':
            throw new ApprovalInvalidError('Token has already been used');
        case 'TARGET_CHANGED':
            throw new TargetChangedError();
    }
    throw new ApprovalInvalidError('Unknown validation error');
}

/** Normalizes validated send params into the payload shape used for hashing/sending. */
function toHashPayload(p: SendEmailParams) {
    return {
        to: p.to,
        subject: p.subject,
        body: p.body,
        bodyType: p.body_type,
        ...(p.cc != null && { cc: p.cc }),
        ...(p.bcc != null && { bcc: p.bcc }),
        ...(p.reply_to != null && { replyTo: p.reply_to }),
        ...(p.attachments != null && { attachments: p.attachments }),
        ...(p.inline_images != null && {
            inlineImages: p.inline_images.map((i) => ({ path: i.path, contentId: i.content_id })),
        }),
        ...(p.account_id != null && { accountId: p.account_id }),
    };
}

/** Builds the AppleScript sender params from validated send params. */
function toMailSenderParams(p: SendEmailParams): MailSenderParams {
    return {
        to: p.to,
        subject: p.subject,
        body: p.body,
        bodyType: p.body_type,
        ...(p.cc != null && { cc: p.cc }),
        ...(p.bcc != null && { bcc: p.bcc }),
        ...(p.reply_to != null && { replyTo: p.reply_to }),
        ...(p.attachments != null && { attachments: p.attachments }),
        ...(p.inline_images != null && {
            inlineImages: p.inline_images.map((img) => ({ path: img.path, contentId: img.content_id })),
        }),
        ...(p.account_id != null && { accountId: p.account_id }),
    };
}

// ---------------------------------------------------------------------------
// OutgoingMailTools — two-phase (prepare/confirm) email sending
// ---------------------------------------------------------------------------

/**
 * Guards outbound email behind a two-phase approval, mirroring the destructive
 * mailbox operations. `prepare_send_email` validates the payload (including
 * attachment path safety), previews it, and returns a single-use token;
 * `confirm_send_email` sends only that exact previewed message.
 *
 * Note: this stops the model from sending in a single unreviewed step, but is
 * only a real safety boundary if the MCP client surfaces the confirm step to a
 * human — the same model can otherwise call both. The path confinement applied
 * here is model-independent.
 */
export class OutgoingMailTools {
    private readonly mailSender: IMailSender;
    private readonly tokenManager: ApprovalTokenManager;

    constructor(mailSender: IMailSender, tokenManager: ApprovalTokenManager) {
        this.mailSender = mailSender;
        this.tokenManager = tokenManager;
    }

    /** Phase 1: validate + preview an outgoing email and issue an approval token. */
    prepareSendEmail(params: SendEmailParams) {
        // Reject sensitive-file exfiltration before a token is ever issued.
        for (const att of params.attachments ?? []) {
            assertSafeAttachmentPath(att.path);
        }
        for (const img of params.inline_images ?? []) {
            assertSafeAttachmentPath(img.path);
        }
        const hash = hashSendPayloadForApproval(toHashPayload(params));
        const token = this.tokenManager.generateToken({
            operation: 'send_email',
            targetType: 'outgoing',
            targetId: SEND_TARGET_ID,
            targetHash: hash,
            metadata: { payload: params },
        });
        const bodyPreview = params.body.length > BODY_PREVIEW_LIMIT
            ? `${params.body.slice(0, BODY_PREVIEW_LIMIT)}…`
            : params.body;
        return {
            token_id: token.tokenId,
            expires_at: new Date(token.expiresAt).toISOString(),
            email: {
                to: params.to,
                cc: params.cc ?? [],
                bcc: params.bcc ?? [],
                subject: params.subject,
                body_type: params.body_type,
                body_preview: bodyPreview,
                attachments: (params.attachments ?? []).map((a) => a.name ?? a.path),
                inline_image_count: params.inline_images?.length ?? 0,
                account_id: params.account_id ?? null,
            },
            action: 'This email will be sent immediately when confirmed with confirm_send_email.',
        };
    }

    /** Phase 2: send the previously prepared email identified by its token. */
    confirmSendEmail(params: ConfirmSendEmailParams) {
        const result = this.tokenManager.consumeToken(params.token_id, 'send_email', SEND_TARGET_ID);
        if (!result.valid) {
            throwValidationError(result.error);
        }
        const token = result.token!;
        const payload = token.metadata['payload'] as SendEmailParams;
        // Integrity check: the sent message must match what was previewed.
        if (hashSendPayloadForApproval(toHashPayload(payload)) !== token.targetHash) {
            throw new TargetChangedError();
        }
        const sent = this.mailSender.sendEmail(toMailSenderParams(payload));
        logAudit('send_email', {
            to: payload.to,
            cc: payload.cc ?? [],
            bcc: payload.bcc ?? [],
            subject: payload.subject,
            accountId: payload.account_id ?? null,
            messageId: sent.messageId,
        });
        return {
            message_id: sent.messageId,
            sent_at: sent.sentAt,
            status: 'sent' as const,
        };
    }
}

/** Creates an OutgoingMailTools instance with the given sender and token manager. */
export function createOutgoingMailTools(mailSender: IMailSender, tokenManager: ApprovalTokenManager): OutgoingMailTools {
    return new OutgoingMailTools(mailSender, tokenManager);
}
