import { createHash } from 'node:crypto';
/**
 * Produces a truncated SHA-256 fingerprint of an email's key properties.
 * Used to detect modifications between the prepare and confirm steps.
 */
export function hashEmailForApproval(email) {
    return createHash('sha256')
        .update(`${email.id}:${email.subject ?? ''}:${email.folderId}:${email.timeReceived ?? 0}`)
        .digest('hex')
        .slice(0, 16);
}
/**
 * Produces a truncated SHA-256 fingerprint of a folder's key properties.
 * Used to detect modifications between the prepare and confirm steps.
 */
export function hashFolderForApproval(folder) {
    return createHash('sha256')
        .update(`${folder.id}:${folder.name ?? ''}:${folder.messageCount}`)
        .digest('hex')
        .slice(0, 16);
}
/**
 * Produces a truncated SHA-256 fingerprint of an outgoing-email payload.
 * Binds a send-email approval token to the exact message the model asked to
 * send, so the confirmed message cannot differ from the previewed one.
 */
export function hashSendPayloadForApproval(payload) {
    const canonical = JSON.stringify({
        to: payload.to,
        cc: payload.cc ?? [],
        bcc: payload.bcc ?? [],
        subject: payload.subject,
        body: payload.body,
        bodyType: payload.bodyType,
        replyTo: payload.replyTo ?? '',
        attachments: (payload.attachments ?? []).map((a) => [a.path, a.name ?? '']),
        inlineImages: (payload.inlineImages ?? []).map((i) => [i.path, i.contentId]),
        accountId: payload.accountId ?? 0,
    });
    return createHash('sha256').update(canonical).digest('hex').slice(0, 16);
}
/**
 * Produces a truncated SHA-256 fingerprint of a calendar event's key
 * properties. Used to detect modifications between the prepare and confirm
 * steps for update/delete. Only fields available on the DB `EventRow` are
 * used (there is no subject at that layer).
 */
export function hashEventForApproval(event) {
    return createHash('sha256')
        .update(`${event.id}:${event.startDate ?? 0}:${event.endDate ?? 0}:${event.uid ?? ''}`)
        .digest('hex')
        .slice(0, 16);
}
