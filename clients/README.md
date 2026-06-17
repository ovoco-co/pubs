# Client Implementations

Each subdirectory contains template tools customized for a specific client's document standards.

## Current Clients

- `nist/` - NIST (National Institute of Standards and Technology)
  - Shared template tools: theme, style architecture, page layout, headers/footers, table formatting
  - `nist/ncstar/` - NCSTAR report assembly and validation
  - `nist/nwirp/` - NWIRP building blocks and snippet documents
- `navy/` - Navy (placeholder, tools in development)

## Adding a New Client

1. Create a folder here named for the client (lowercase, e.g., `jpl/`)
2. Use `nist/` as a reference for the shared template tool pattern
3. Copy and rename the `Set-NIST*` scripts to `Set-<Client>*`
4. Modify color palettes, style names, page margins, header/footer layouts
5. Add document-set-specific subfolders as needed (following the `nist/ncstar/` and `nist/nwirp/` pattern)
6. Add a README.md describing the client's template architecture

## Rules

- Client scripts carry the client prefix in the noun: `Set-NISTTheme`, `Set-NavyTheme`
- Client scripts do not import or reference other clients
- Generic tools in `../word/` are never modified for a specific client
- Each client's tools are self-contained within their directory
