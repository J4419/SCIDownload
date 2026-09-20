# Security policy

## Reporting a vulnerability

Please use GitHub's private vulnerability reporting feature when available. Do not open a public issue containing credentials, institution login details, session cookies, signed download URLs, or an affected browser profile. If private reporting is unavailable, open a minimal issue asking the maintainers for a private contact channel.

Include the affected version, operating system, a minimal reproduction, and the security impact. Remove DOI lists, institution names, proxy credentials, cookies, access tokens, and local paths unless they are essential and safe to disclose.

## Sensitive local data

The dedicated browser profile may contain authenticated session data. Treat it like a browser account: do not commit it, include it in a release archive, upload it for debugging, or use `--keep-browser` on a shared computer. Close the browser after use.

`run.log`, `manifest.jsonl`, screenshots, and copied error messages may expose research interests, DOI lists, publisher/institution context, and local state. Version 1.5 redacts the main program's public IP, absolute local paths, and URL query strings by default, but users should still review artifacts before sharing them.

The CDP debugging endpoint is bound to `127.0.0.1` and startup fails if the requested port is already occupied. Do not modify it to listen on a public interface. `1.check_exit_ip.py` masks the public IP by default; `--show-full-ip` should not be used in screenshots intended for public posting.

## Cache cleanup

`4.ClearCache.py` normally removes cache directories while preserving login state. `--all` removes the complete dedicated profile. A custom `--profile` is rejected for destructive cleanup unless `--confirm-custom-profile` is also supplied. Always use `--dry-run` first when the target is unfamiliar.

## Scope

This project automates access the user already possesses. Security reports should not include methods intended to evade publisher access controls, rate limits, verification systems, or subscription requirements.
