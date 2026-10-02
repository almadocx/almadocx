/** Security limits for untrusted DOCX/ODT packages. */
export const PACKAGE_LIMITS = {
  maxEntries: 5_000,
  maxCompressedBytes: 50 * 1024 * 1024,
  maxUncompressedEntryBytes: 50 * 1024 * 1024,
  maxTotalUncompressedBytes: 100 * 1024 * 1024,
  /** Reject entries where uncompressed/compressed exceeds this (zip bomb). */
  maxCompressionRatio: 100,
  maxXmlBytes: 20 * 1024 * 1024,
  maxPathLength: 512,
} as const
