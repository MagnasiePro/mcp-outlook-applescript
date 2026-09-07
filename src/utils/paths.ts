/**
 * Filesystem path safety checks for outbound/download operations.
 *
 * The MCP tools let a model choose arbitrary filesystem paths: `save_path`
 * for downloaded attachments (arbitrary write → possible persistence) and
 * `attachments[].path` for outgoing mail (arbitrary read → possible
 * exfiltration of sensitive files). These helpers confine those paths.
 *
 * Configuration (environment variables):
 *  - OUTLOOK_MCP_DOWNLOAD_DIR: base directory downloads must stay within.
 *    Defaults to ~/Downloads.
 *  - OUTLOOK_MCP_ATTACHMENT_DIR: if set, outgoing attachments must stay
 *    within this directory (strict allowlist mode). If unset, a denylist of
 *    sensitive locations is enforced instead (so ordinary files still work).
 */
import { homedir } from 'node:os';
import { isAbsolute, resolve, sep } from 'node:path';
import { ValidationError } from './errors.js';

/** Expands a leading `~` (or `~/`) to the current user's home directory. */
export function expandHome(p: string): string {
    if (p === '~') {
        return homedir();
    }
    if (p.startsWith('~/')) {
        return homedir() + p.slice(1);
    }
    return p;
}

/** Resolves a path to an absolute, normalized form with `~` expansion. */
function normalizeAbsolute(p: string): string {
    const expanded = expandHome(p);
    if (!isAbsolute(expanded)) {
        throw new ValidationError(`Path must be absolute: ${p}`);
    }
    return resolve(expanded);
}

/** Returns true when `child` is `base` itself or nested inside `base`. */
function isWithin(base: string, child: string): boolean {
    const normalizedBase = resolve(base);
    return child === normalizedBase || child.startsWith(normalizedBase + sep);
}

/**
 * Substrings and basenames that must never be an attachment source.
 * Matched case-insensitively against the resolved absolute path.
 */
const SENSITIVE_PATH_SEGMENTS = [
    `${sep}.ssh${sep}`,
    `${sep}.aws${sep}`,
    `${sep}.gnupg${sep}`,
    `${sep}.gcloud${sep}`,
    `${sep}.kube${sep}`,
    `${sep}.docker${sep}`,
    `${sep}.config${sep}`,
    `${sep}library${sep}keychains${sep}`,
    `${sep}etc${sep}`,
];

/** Basename patterns that indicate secret material. */
const SENSITIVE_BASENAMES = [
    'id_rsa', 'id_dsa', 'id_ecdsa', 'id_ed25519',
    '.netrc', '.npmrc', '.pgpass', '.htpasswd',
];

/** File extensions that typically hold private keys / credentials. */
const SENSITIVE_EXTENSIONS = ['.pem', '.key', '.p12', '.pfx', '.keychain'];

/**
 * Validates a destination path for a downloaded attachment.
 * Confines writes to OUTLOOK_MCP_DOWNLOAD_DIR (default ~/Downloads) and
 * rejects path traversal outside it.
 * @returns The resolved absolute path to write to.
 * @throws ValidationError if the path escapes the allowed base directory.
 */
export function assertSafeDownloadPath(savePath: string): string {
    const base = expandHome(process.env['OUTLOOK_MCP_DOWNLOAD_DIR'] ?? `${homedir()}${sep}Downloads`);
    const resolved = normalizeAbsolute(savePath);
    if (!isWithin(base, resolved)) {
        throw new ValidationError(
            `Refusing to write outside the allowed download directory (${resolve(base)}). ` +
            `Set OUTLOOK_MCP_DOWNLOAD_DIR to change it.`,
        );
    }
    return resolved;
}

/**
 * Validates a source path for an outgoing-mail attachment or inline image.
 * In allowlist mode (OUTLOOK_MCP_ATTACHMENT_DIR set) the path must stay
 * inside that directory. Otherwise a denylist of sensitive locations
 * (SSH/GPG/cloud creds, keychains, private keys, /etc, dotfile config) is
 * rejected to blunt file-exfiltration via email.
 * @returns The resolved absolute path.
 * @throws ValidationError if the path is disallowed.
 */
export function assertSafeAttachmentPath(attachmentPath: string): string {
    const resolved = normalizeAbsolute(attachmentPath);
    const allowDir = process.env['OUTLOOK_MCP_ATTACHMENT_DIR'];
    if (allowDir != null && allowDir.length > 0) {
        const base = expandHome(allowDir);
        if (!isWithin(base, resolved)) {
            throw new ValidationError(
                `Attachment path must be inside the allowed attachment directory (${resolve(base)}).`,
            );
        }
        return resolved;
    }
    const lower = resolved.toLowerCase();
    const basename = lower.split(sep).pop() ?? '';
    const disallowed =
        SENSITIVE_PATH_SEGMENTS.some((seg) => lower.includes(seg)) ||
        SENSITIVE_BASENAMES.includes(basename) ||
        SENSITIVE_EXTENSIONS.some((ext) => basename.endsWith(ext));
    if (disallowed) {
        throw new ValidationError(
            `Refusing to attach a sensitive file: ${attachmentPath}. ` +
            `Set OUTLOOK_MCP_ATTACHMENT_DIR to restrict attachments to a specific directory.`,
        );
    }
    return resolved;
}
