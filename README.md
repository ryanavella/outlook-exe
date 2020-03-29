# outlook-exe
-------------

Convenience wrappers for command-line invocation of Outlook.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://github.com/ryanavella/outlook-exe/blob/master/LICENSE-MIT) [![License: Unlicense](https://img.shields.io/badge/license-Unlicense-blue.svg)](https://github.com/ryanavella/outlook-exe/blob/master/LICENSE-UNLICENSE)

Dual-licensed under MIT or the [UNLICENSE](https://unlicense.org).

## Example

Basic usage:

```rust
use outlook_exe;

outlook_exe::MessageBuilder::new()
    .with_recipient("noreply@example.org")
    .with_subject("Hello, World!")
    .with_body("Line with spaces\nAnother line")
    .with_attachment("C:/tmp/file.txt")
    .spawn()
    .unwrap();
```
