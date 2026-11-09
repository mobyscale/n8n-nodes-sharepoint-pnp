export { parseItemId } from '../../shared/pickers';

export function toArrayBuffer(buffer: Buffer): ArrayBuffer {
	return buffer.buffer.slice(
		buffer.byteOffset,
		buffer.byteOffset + buffer.byteLength,
	) as ArrayBuffer;
}

/** Document libraries have base type 1. Only other lists can have attachments. */
export function getAttachmentSupportError(info: {
	BaseType?: number;
	EnableAttachments?: boolean;
}): string | undefined {
	if (info.BaseType === 1) {
		return 'Document libraries do not support attachments. Use the File resource for files in a library.';
	}
	if (info.EnableAttachments === false) {
		return 'Attachments are turned off for this list. Turn on attachments in the list settings first.';
	}
	return undefined;
}
