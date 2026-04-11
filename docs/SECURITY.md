# Security

This document describes how teams-cli handles authentication, credentials, and data in transit.

## Authentication

teams-cli authenticates via a Playwright-based browser login flow — the same sign-in experience as the Teams web client, including full MFA support. No credentials are ever entered into or handled by the CLI directly; authentication is delegated entirely to Microsoft's login page in a real browser.

A persistent browser profile preserves session cookies so that expired tokens can be refreshed automatically via a headless browser, without requiring user interaction unless the session itself has expired.

Running `teams logout` removes all stored tokens and the browser profile directory.

## Credential storage

- Tokens are stored locally at `~/.teams-cli/tokens.json` and are never logged, printed, or transmitted to any service other than Microsoft's APIs
- The `~/.teams-cli/` directory is gitignored to prevent accidental commits to version control
- Token lifetimes are governed by the organisation's Entra ID policy, limiting the window of validity

## Transport security

- All API communication uses HTTPS with certificate validation via Node.js native `fetch`
- Bearer tokens are transmitted only in the `Authorization` header over TLS
- No sensitive data is sent to third-party services — all requests go directly to Microsoft endpoints

## Input handling

- API request bodies are constructed using structured JSON serialization, preventing malformed payloads
- URL parameters are encoded via `URLSearchParams`
- Server-side validation is performed by Microsoft's APIs on all inputs

## Dependencies

All dependencies are well-maintained, widely used packages pinned via `package-lock.json`:

- **commander** — CLI framework (no network access)
- **chalk** — terminal colour output (display only)
- **cli-table3** — terminal table rendering (display only)
- **playwright** — browser automation (used only for authentication, dynamically imported at runtime)

The repository has the following GitHub security features enabled:

- **Dependabot security updates** — vulnerable dependencies are flagged and patched automatically
- **Secret scanning** with **push protection** — prevents accidental commits of tokens or credentials
- **Code scanning** — static analysis via GitHub's CodeQL to detect security issues in code

## Package publishing

- Releases are fully automated via [semantic-release](https://github.com/semantic-release/semantic-release) — version bumps, changelogs, and npm publishing are determined by conventional commit messages, removing human error from the release process
- npm publishing uses **OIDC trusted publishing** — the GitHub Actions workflow authenticates directly with npm via OpenID Connect, eliminating the need for long-lived npm access tokens
- Published packages include **provenance attestations**, allowing consumers to verify that a package was built from this repository's CI pipeline and has not been tampered with
- No npm tokens or publishing credentials are stored as repository secrets

## Reporting security issues

If you discover a security issue, please open an issue on the [GitHub repository](https://github.com/mingnz/teams-cli/issues) or contact the maintainer directly.
