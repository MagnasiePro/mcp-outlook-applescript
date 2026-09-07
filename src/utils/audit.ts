/**
 * Lightweight audit logging for sensitive actions (send, delete, move,
 * download, calendar writes).
 *
 * Lines are written to STDERR as JSON — never stdout, which carries the
 * MCP JSON-RPC stream and must not be polluted. This gives an operator a
 * forensic trail of outbound/destructive operations the model performed.
 */

/** Emits one structured audit line to stderr. Never throws. */
export function logAudit(action: string, details: Record<string, unknown> = {}): void {
    try {
        const entry = {
            ts: new Date().toISOString(),
            audit: action,
            ...details,
        };
        process.stderr.write(`${JSON.stringify(entry)}\n`);
    }
    catch {
        // Auditing must never break the operation it is recording.
    }
}
