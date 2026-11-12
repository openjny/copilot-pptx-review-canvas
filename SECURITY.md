# Security policy

## Supported versions

Security fixes go into the latest release. Please update before reporting.

## Reporting a vulnerability

Please do not open a public issue for a security problem. Use GitHub's private reporting instead: open the **Security** tab of this repository and choose **Report a vulnerability**. Include the version, the steps to reproduce, and what an attacker could do.

You can expect an acknowledgement within a few days. This is a spare-time project, so a fix may take longer; I will keep you updated and credit you in the release notes if you wish.

## What the extension does, for scoping a report

- It starts a small HTTP server on `127.0.0.1` with a random port for each open canvas. Requests are accepted only for the loopback host, and `POST` requests must be JSON from the same origin. Anything that lets another web page or another user on the machine drive this server is in scope.
- It opens `.pptx` files read-only through desktop PowerPoint (COM automation) with macros disabled and without prompts, and renders them to PNG.
- It rewrites a `.pptx` file only when you reorder or delete slides, after keeping a backup under `~/.copilot/extension-data/pptx-review-canvas/backups/`. Anything that makes it write outside the chosen file, or that corrupts a deck, is in scope.
- It stores marks, comments and rendered images locally under `~/.copilot/extension-data/pptx-review-canvas/` and makes no network requests of its own: no telemetry, no remote assets.

Vulnerabilities in PowerPoint, VS Code, the Copilot clients or Node.js belong to their own projects.
