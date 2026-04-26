# Security Policy

## Supported Versions

We release patches for security vulnerabilities. Currently supported versions:

| Version | Supported          |
| ------- | ------------------ |
| 0.1.x   | :white_check_mark: |

## Reporting a Vulnerability

We take the security of office_oxide seriously. If you believe you have found a security vulnerability, please report it to us as described below.

### Where to Report

**Please do not report security vulnerabilities through public GitHub issues.**

Instead, please email security reports directly to **yfedoseev@gmail.com** with the subject line `[office_oxide] Security Vulnerability Report`.

### What to Include

Please include the following information in your report:

* Type of issue (e.g. buffer overflow, XXE injection, zip bomb, etc.)
* Full paths of source file(s) related to the manifestation of the issue
* The location of the affected source code (tag/branch/commit or direct URL)
* Any special configuration required to reproduce the issue
* Step-by-step instructions to reproduce the issue
* Proof-of-concept or exploit code (if possible)
* Impact of the issue, including how an attacker might exploit it

### What to Expect

* We will acknowledge your email within 48 hours
* We will send a more detailed response within 7 days indicating the next steps
* We will keep you informed about progress towards a fix
* We may ask for additional information or guidance
* Once fixed, we will publicly disclose the vulnerability (crediting you if desired)

## Office Document Security Considerations

Office documents can contain potentially malicious content. This library:

* **Validates input**: All document inputs are validated for structure and size limits
* **Limits recursion**: Maximum recursion depth prevents stack overflow
* **Resource limits**: Maximum file size, object count, and memory usage limits
* **Safe parsing**: No unsafe code in critical parsing paths
* **Sandboxing recommended**: For processing untrusted documents, run in a sandboxed environment

### Known Risks

When processing untrusted Office documents:

1. **Zip bombs**: OOXML documents are ZIP files that can contain highly compressed content
   - Mitigation: Size limits on decompressed streams, CRC tolerance

2. **XML bombs (billion laughs)**: Malformed XML with entity expansion
   - Mitigation: quick-xml does not expand entities by default

3. **Resource exhaustion**: Large or complex documents can consume significant CPU/memory
   - Mitigation: Shared string DoS cap (32,768 chars), configurable resource limits

4. **Malformed documents**: Crafted documents may trigger edge cases
   - Mitigation: Extensive validation, tolerant parsing with graceful degradation

5. **OLE2 containers (legacy formats)**: CFB files can have circular references or corrupted FAT chains
   - Mitigation: Visited-sector tracking, chain length limits

### Best Practices

When using office_oxide with untrusted documents:

1. **Timeout operations**: Use timeouts for document processing
2. **Sandbox execution**: Run in containers or VMs when processing untrusted files
3. **Validate sources**: Only process documents from trusted sources when possible
4. **Monitor resources**: Track memory and CPU usage
5. **Update regularly**: Keep office_oxide updated with latest security patches

## Disclosure Policy

When we receive a security bug report, we will:

1. Confirm the problem and determine affected versions
2. Audit code to find similar problems
3. Prepare fixes for all supported versions
4. Release patches as soon as possible

We ask security researchers to:

* Give us reasonable time to respond before public disclosure
* Make a good faith effort to avoid privacy violations and service disruption
* Not access or modify other users' data

## Comments on this Policy

If you have suggestions on how this process could be improved, please submit a pull request.
