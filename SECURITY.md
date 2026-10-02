# Almadocx Security

Document packages (DOCX/ODT) are treated as **untrusted input**.

## Zip

- Reject absolute paths and `..` segments in entry names
- Cap entry count, compressed size, uncompressed size, and compression ratio
- Bound total uncompressed bytes across the package

## XML

- Disable DTDs and external entities (XXE)
- Cap document size before parse
- Do not resolve network resources during parse

## Macros

- Detect and preserve macro parts when present
- Never execute macros; inert by default until an explicit sandboxed policy exists

## Clipboard

- Sanitize HTML paste; never evaluate script from clipboard content
