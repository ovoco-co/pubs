# FrameMaker Template Tooling

Tools for building and managing Adobe FrameMaker templates. Organized into two paths reflecting FrameMaker's two operating modes.

## Unstructured FrameMaker

Traditional FrameMaker using paragraph formats, character formats, table formats, master pages, and reference pages. Used for long-form technical manuals where content is formatted text, not XML elements.

Tools will include master page generators, paragraph catalog builders, and import format mapping tools.

See `unstructured/` for details.

## Structured FrameMaker

FrameMaker operating with an Element Definition Document (EDD) tied to a DTD or schema. Content is authored as XML elements. Multiple authoring standards are supported:

- `structured/dita/` - DITA (Darwin Information Typing Architecture). Topic-based XML standard. Note: `../../dita/` converts Word content INTO DITA format, while this directory covers publishing DITA content THROUGH FrameMaker. Different workflows.
- `structured/s1000d/` - S1000D. International specification for technical publications, widely used in military and aerospace.
- `structured/docbook/` - DocBook. Book-oriented XML standard for technical documentation.

## Relationship to Other Publishing Paths

The publishing toolkit has three independent paths:

- `../word/` - Generic OOXML inspection and manipulation (any Word document)
- `../dita/` - Converting Word content into DITA XML format
- `../frame/` (this directory) - FrameMaker template tooling for both unstructured and structured authoring

Each path can be used independently. A team doing unstructured FrameMaker work does not need the DITA conversion tools. A team doing DITA authoring in FrameMaker needs `structured/dita/` but not `../dita/` (unless they are also converting legacy Word content).
