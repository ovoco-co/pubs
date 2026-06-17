# Client Implementations

Each subdirectory contains template tools customized for a specific client's document standards.

## Current Clients

- `nist/` - NIST (National Institute of Standards and Technology)
  - Shared template tools: theme, style architecture, page layout, headers/footers, table formatting
  - `nist/ncstar/` - NCSTAR report assembly and validation
  - `nist/nwirp/` - NWIRP building blocks and snippet documents
- `navy/` - Navy (placeholder, tools in development)

## Architecture

Generic OOXML tools live in `../word/`. Client scripts are thin wrappers that supply client-specific configuration (color palettes, style lists, margin defaults) and delegate to the generic tools. This keeps the OOXML logic in one place and prevents duplication across clients.

## Adding a New Client

1. Create a folder here named for the client (lowercase, e.g., `jpl/`)
2. Use `nist/` as a reference for the thin-wrapper pattern
3. Create `Set-<Client>Theme.ps1` calling `../word/Set-ThemeColors.ps1` with the client palette
4. Create `Set-<Client>StyleArchitecture.ps1` importing `../word/OoXml/OoXml.psd1` with client style config
5. Create thin wrappers for page layout, audit, and other generic tools as needed
6. Add document-set-specific subfolders as needed (following the `nist/ncstar/` and `nist/nwirp/` pattern)
7. Add a README.md describing the client's template architecture

## Rules

- Client scripts carry the client prefix in the noun: `Set-NISTTheme`, `Set-NavyTheme`
- Client scripts delegate to generic tools in `../word/`, supplying only client-specific data
- Client scripts do not import or reference other clients
- Generic tools in `../word/` are never modified for a specific client
