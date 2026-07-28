# Client Implementations

Per-client notes. The NIST template tools were consolidated into `../word/` so all the Word scripts live in one folder; see `../word/README.md`.

## Current Clients

- **NIST** (National Institute of Standards and Technology) - tools now live in `../word/`:
  - `../word/Set-NIST*.ps1`, `../word/Audit-NISTStyleSet.ps1` - theme, style architecture, headers/footers, audit
  - `../word/ncstar/` - NCSTAR report assembly and validation
  - `../word/nwirp/` - NWIRP building blocks and snippet documents
- `navy/` - Navy (placeholder, tools in development)

## Architecture

The generic OOXML engines live in `../word/`. Client-specific scripts are thin wrappers that supply client configuration (color palettes, style lists, margin defaults) and delegate to those engines, which keeps the OOXML logic in one place. NIST wrappers sit alongside the engines in `../word/`; a wrapper finds its engine and the `OoXml/` module by its own location, so it must stay in that folder.

## Rules

- Client scripts carry the client prefix in the noun: `Set-NISTTheme`, `Set-NavyTheme`
- Client scripts delegate to the generic engines, supplying only client-specific data
- Client scripts do not import or reference other clients
- The generic engines are never modified for a specific client
