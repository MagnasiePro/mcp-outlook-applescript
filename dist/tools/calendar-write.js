import { z } from 'zod';
import { hashEventForApproval, hashSendPayloadForApproval, } from '../approval/index.js';
import { logAudit } from '../utils/audit.js';
import { CreateEventInput, UpdateEventInput, DeleteEventInput, } from './calendar.js';
import { ApprovalExpiredError, ApprovalInvalidError, TargetChangedError, NotFoundError, } from '../utils/errors.js';
// ---------------------------------------------------------------------------
// Input schemas — two-phase calendar writes
// ---------------------------------------------------------------------------
export const PrepareCreateEventInput = CreateEventInput;
export const ConfirmCreateEventInput = z.strictObject({
    token_id: z.uuid().describe('The approval token UUID from prepare_create_event'),
});
export const PrepareUpdateEventInput = UpdateEventInput;
export const ConfirmUpdateEventInput = z.strictObject({
    token_id: z.uuid().describe('The approval token UUID from prepare_update_event'),
    event_id: z.number().int().positive().describe('The event ID to update — must match the prepare step'),
});
export const PrepareDeleteEventInput = DeleteEventInput;
export const ConfirmDeleteEventInput = z.strictObject({
    token_id: z.uuid().describe('The approval token UUID from prepare_delete_event'),
    event_id: z.number().int().positive().describe('The event ID to delete — must match the prepare step'),
});
/** create_event has no pre-existing entity, so its token uses a synthetic target id. */
const CREATE_TARGET_ID = 0;
/** Maps an approval validation failure reason to a typed error. */
function throwValidationError(reason) {
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
/** Maps Zod-validated CreateEventInput fields to the CalendarWriter's internal param shape. */
export function buildCalendarWriterParams(params) {
    let recurrence;
    if (params.recurrence != null) {
        const rec = params.recurrence;
        recurrence = {
            frequency: rec.frequency,
            interval: rec.interval,
            ...(rec.days_of_week != null && { daysOfWeek: rec.days_of_week }),
            ...(rec.day_of_month != null && { dayOfMonth: rec.day_of_month }),
            ...(rec.week_of_month != null && { weekOfMonth: rec.week_of_month }),
            ...(rec.day_of_week_monthly != null && { dayOfWeekMonthly: rec.day_of_week_monthly }),
            ...(rec.end.type === 'end_date' && { endDate: rec.end.date }),
            ...(rec.end.type === 'end_after_count' && { endAfterCount: rec.end.count }),
        };
    }
    return {
        title: params.title,
        startDate: params.start_date,
        endDate: params.end_date,
        ...(params.calendar_id != null && { calendarId: params.calendar_id }),
        ...(params.location != null && { location: params.location }),
        ...(params.description != null && { description: params.description }),
        ...(params.is_all_day != null && { isAllDay: params.is_all_day }),
        ...(recurrence != null && { recurrence }),
    };
}
/** Builds the EventUpdates delta from validated update params. */
function buildEventUpdates(params) {
    return {
        ...(params.title != null && { title: params.title }),
        ...(params.start_date != null && { startDate: params.start_date }),
        ...(params.end_date != null && { endDate: params.end_date }),
        ...(params.location != null && { location: params.location }),
        ...(params.description != null && { description: params.description }),
        ...(params.is_all_day != null && { isAllDay: params.is_all_day }),
    };
}
// ---------------------------------------------------------------------------
// CalendarWriteApprovalTools — two-phase (prepare/confirm) calendar writes
// ---------------------------------------------------------------------------
/**
 * Guards calendar create/update/delete behind a two-phase approval, mirroring
 * the destructive mailbox operations. Update/delete tokens are bound to the
 * event's hash and re-verified at confirm (detects the event changing between
 * steps); create tokens are bound to a hash of the intended payload.
 *
 * As with all approvals here, this is a real boundary only if the MCP client
 * surfaces the confirm step to a human.
 */
export class CalendarWriteApprovalTools {
    repository;
    calendarWriter;
    calendarManager;
    tokenManager;
    constructor(repository, calendarWriter, calendarManager, tokenManager) {
        this.repository = repository;
        this.calendarWriter = calendarWriter;
        this.calendarManager = calendarManager;
        this.tokenManager = tokenManager;
    }
    // --- Create ---
    prepareCreateEvent(params) {
        const hash = hashSendPayloadForApproval({
            to: [],
            subject: params.title,
            body: `${params.start_date}|${params.end_date}|${params.location ?? ''}|${params.description ?? ''}`,
            bodyType: params.is_all_day ? 'allday' : 'timed',
        });
        const token = this.tokenManager.generateToken({
            operation: 'create_event',
            targetType: 'outgoing',
            targetId: CREATE_TARGET_ID,
            targetHash: hash,
            metadata: { payload: params },
        });
        return {
            token_id: token.tokenId,
            expires_at: new Date(token.expiresAt).toISOString(),
            event: {
                title: params.title,
                start_date: params.start_date,
                end_date: params.end_date,
                calendar_id: params.calendar_id ?? null,
                location: params.location ?? null,
                is_all_day: params.is_all_day,
                is_recurring: params.recurrence != null,
            },
            action: 'This event will be created when confirmed with confirm_create_event.',
        };
    }
    confirmCreateEvent(params) {
        const result = this.tokenManager.consumeToken(params.token_id, 'create_event', CREATE_TARGET_ID);
        if (!result.valid) {
            throwValidationError(result.error);
        }
        const payload = result.token.metadata['payload'];
        const created = this.calendarWriter.createEvent(buildCalendarWriterParams(payload));
        logAudit('create_event', { title: payload.title, eventId: created.id, calendarId: created.calendarId });
        return {
            id: created.id,
            title: payload.title,
            start_date: payload.start_date,
            end_date: payload.end_date,
            calendar_id: created.calendarId,
            location: payload.location ?? null,
            description: payload.description ?? null,
            is_all_day: payload.is_all_day,
            is_recurring: payload.recurrence != null,
        };
    }
    // --- Update ---
    prepareUpdateEvent(params) {
        const token = this.tokenManager.generateToken({
            operation: 'update_event',
            targetType: 'event',
            targetId: params.event_id,
            targetHash: this.requireEventHash(params.event_id),
            metadata: { updates: params },
        });
        return {
            token_id: token.tokenId,
            expires_at: new Date(token.expiresAt).toISOString(),
            event_id: params.event_id,
            apply_to: params.apply_to,
            updated_fields: Object.keys(buildEventUpdates(params)),
            action: 'This event will be updated when confirmed with confirm_update_event.',
        };
    }
    confirmUpdateEvent(params) {
        const token = this.consumeAndVerifyEvent(params.token_id, 'update_event', params.event_id);
        const updateParams = token.metadata['updates'];
        const result = this.calendarManager.updateEvent(params.event_id, buildEventUpdates(updateParams), updateParams.apply_to);
        logAudit('update_event', { eventId: params.event_id, applyTo: updateParams.apply_to, updatedFields: result.updatedFields });
        return {
            success: true,
            message: `Event ${result.id} updated. Updated fields: ${result.updatedFields.join(', ')}`,
        };
    }
    // --- Delete ---
    prepareDeleteEvent(params) {
        const token = this.tokenManager.generateToken({
            operation: 'delete_event',
            targetType: 'event',
            targetId: params.event_id,
            targetHash: this.requireEventHash(params.event_id),
            metadata: { applyTo: params.apply_to },
        });
        return {
            token_id: token.tokenId,
            expires_at: new Date(token.expiresAt).toISOString(),
            event_id: params.event_id,
            apply_to: params.apply_to,
            action: `This event${params.apply_to === 'all_in_series' ? ' (entire series)' : ''} will be permanently deleted when confirmed with confirm_delete_event.`,
        };
    }
    confirmDeleteEvent(params) {
        const token = this.consumeAndVerifyEvent(params.token_id, 'delete_event', params.event_id);
        const applyTo = token.metadata['applyTo'];
        this.calendarManager.deleteEvent(params.event_id, applyTo);
        logAudit('delete_event', { eventId: params.event_id, applyTo });
        const seriesText = applyTo === 'all_in_series' ? ' (entire series)' : '';
        return { success: true, message: `Event ${params.event_id}${seriesText} deleted.` };
    }
    // --- Helpers ---
    requireEventHash(eventId) {
        const event = this.repository.getEvent(eventId);
        if (event == null) {
            throw new NotFoundError('Event', eventId);
        }
        return hashEventForApproval(event);
    }
    consumeAndVerifyEvent(tokenId, operation, eventId) {
        const result = this.tokenManager.consumeToken(tokenId, operation, eventId);
        if (!result.valid) {
            throwValidationError(result.error);
        }
        const token = result.token;
        const event = this.repository.getEvent(eventId);
        if (event == null) {
            throw new NotFoundError('Event', eventId);
        }
        if (hashEventForApproval(event) !== token.targetHash) {
            throw new TargetChangedError();
        }
        return token;
    }
}
/** Creates a CalendarWriteApprovalTools instance. */
export function createCalendarWriteApprovalTools(repository, calendarWriter, calendarManager, tokenManager) {
    return new CalendarWriteApprovalTools(repository, calendarWriter, calendarManager, tokenManager);
}
