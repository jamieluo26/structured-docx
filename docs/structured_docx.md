# Structured editable DOCX, v1

This offline byte API converts a finite typed document into editable Word text,
headings and native lists. It independently reads the actual selected DOCX and
compares the complete ordered supported content with the supplied source. It
does not read input paths, save output files, discover projects or access a
network. The caller owns output storage and visual QA.

The generation dependency is `python-docx`; verification uses Python's standard
ZIP/XML libraries. Dependencies and fonts are not bundled or installed by this
tool. No CLI runs automatically. Importing the synthetic example does not
generate a document.

## Source model

Supply explicit UTF-8 JSON bytes:

```json
{
  "version": "structured-docx/v1",
  "title": "Editable document",
  "blocks": [
    {"type": "heading", "level": 1, "text": "A heading"},
    {"type": "paragraph", "text": "  Literal text\tand a tab\nthen a line.  "},
    {"type": "list", "ordered": true, "items": [
      {"text": "First", "children": {"type": "list", "ordered": false,
        "items": [{"text": "Nested bullet"}]}},
      {"text": "Second"}
    ]}
  ]
}
```

Only the shown fields plus optional root `author`, `language` and `createdAt`
are accepted. Heading levels are exact integers 1–3; booleans are not levels.
`ordered` must be an exact boolean. Text must be a string. Unknown fields,
duplicate JSON keys, non-finite numbers, trailing input, invalid UTF-8 and
XML-invalid control characters/surrogates are rejected before generation.

Empty title, paragraphs and list-item text are allowed and remain editable.
The body-block array may be empty. Every list needs at least one item; an empty
list/empty child list is rejected as `LIST_EMPTY`. Each item may contain exactly
one child list, to four levels. Mixed ordered/bullet nesting is supported. Every
list instance has a distinct native numbering ID and starts at one; a parent
list continues after a child list. Literal numeric prefixes are never used as
the list mechanism.

CRLF and CR project to LF. All remaining whitespace and Unicode are preserved;
there is no trimming, Markdown/HTML interpretation, implicit label creation,
placeholder, arbitrary object flattening or truncation. Tabs become native tab
elements and LF becomes native line breaks. Leading/trailing whitespace in
text runs is explicitly preserved.

Optional `author` is at most 255 characters. `language` is a 2–8 letter primary
subtag followed by optional 1–8 alphanumeric subtags (maximum 63 characters).
`createdAt` is a valid UTC `YYYY-MM-DDTHH:MM:SSZ` timestamp. Defaults are empty
author, `en-US`, and `2000-01-01T00:00:00Z`. Core title/subject/description and
other optional metadata are neutral empty values; the full source title is
the first editable Title paragraph. No product/company metadata is inherited.

## API

```python
rendered = render_docx(source_bytes, expected_source_sha256=None, limits=None)
checked = verify_docx(
    source_bytes,
    rendered.docx_bytes,
    expected_docx_sha256=rendered.docx_sha256,
    expected_render_profile_sha256=rendered.render_profile_sha256,
    limits=None,
)
```

`RenderResult` has `docx_bytes`, `mime_type`, `filename`, `source_sha256`,
`model_sha256`, `render_profile_sha256`, `docx_sha256`, `docx_size`, and `receipt`.
The proposed filename is always `structured-document.docx`; no caller text is
turned into a filesystem path. `DocumentError` exposes stable `code`, `path`,
and `stage`; its message contains no source body text or private filesystem path.
Generation errors raise it without a valid artifact result.

`Verification` has `status`, `code`, those five SHA/size fields, `mismatches`,
`unknowns`, `observed`, and `receipt`. Its `observed` items are actual ordered
OOXML text/structure records. Mismatches contain paths and text hashes rather
than raw text. Supported results are `PASS`, `FAIL`, and `NOT_VERIFIED`.
`PASS` requires complete supported semantic/binding checks and both trusted
expected output/profile SHA identities. Passing `None` for an identity gives
a limited check with `MISSING_EXPECTED_IDENTITY`, never full `PASS`.
Unexpected verification/dependency problems are honest unknown coverage.

Caller-selected identity must come from a trusted frozen render receipt, not
from hashing an arbitrary file and then treating that hash as provenance.
Wrong output SHA and any established source/body mismatch remain failures
even if another layer cannot be inspected. A matching embedded summary alone
cannot make changed, missing, duplicated or reordered content pass.

## Actual package and readback

Generation uses the maintained document library for the actual body/runs/native
list references. It retains only that generated body and supplies a finite
neutral package: main document, styles, numbering, core/app properties, source
binding JSON, content types and the two required relationship parts. It uses
local default-namespace serialization for the OPC content-type and relationship
roots, without changing a process-global namespace registry. It uses
Letter portrait, one-inch margins, black Arial 11 pt body text with Microsoft
YaHei as the East Asian font, and black 22/16/14/12 pt title/headings. Installed
glyph availability and rendered layout require separate real rendering.

Verification independently walks the ZIP, XML body, runs, whitespace elements,
styles, page setup, metadata, relationships and numbering definitions. Its
expected source inventory uses a separate iterative traversal; generation does
not use that verifier projection. It checks list kind, four-level native number
references, start-at-one overrides, per-instance numbering ownership and item
order. It compares every supported item, including the full final text tail.

The binding records raw source SHA, canonical normalized-model SHA and render
profile SHA. Raw-byte identity is separate from CR/LF semantic projection.
These markers establish consistency, not signatures, authorship or rights.
Receipts also record actual output SHA and runtime/dependency provenance.
Semantic repeatability is separate from byte reproducibility across library
and Python versions.

Only the finite package members and internal relationship targets are accepted.
An optional bounded JPEG package thumbnail is allowed as an inert profiled part,
with the corresponding content type and thumbnail relationship. Tables, images
in the document, links, macros, headers/footers, fields, textboxes, tracked
changes, comments, embedded objects, external relationships and unknown parts
are unsupported. DTD/entities, malformed XML, encrypted/unsupported ZIP methods,
duplicate/traversal paths, ZIP64, archive comments, prepended/appended payloads,
member padding and resource-limit violations cannot yield fidelity `PASS`.

## Fixed bounds

`limits` is an optional dictionary of exact positive integers. Keys below may
only lower these hard maxima; unknown keys, bool values and raised limits fail.

| Key | Maximum |
| --- | ---: |
| `source_bytes` | 2,097,152 |
| `text_chars` | 500,000 code points across supplied text fields |
| `string_chars` | 32,768 code points |
| `blocks` | 1,000 |
| `list_items` | 2,000 |
| `list_depth` | 4 |
| `model_depth` | 16 JSON container levels |
| `archive_bytes` | 8,388,608 |
| `zip_members` | 128 |
| `uncompressed_bytes` | 33,554,432 |
| `xml_bytes` | 4,194,304 per XML part |
| `compression_ratio` | 1,000 |

XML work also has a fixed 64-level/200,000-element per-part ceiling. At most
128 detailed mismatches are returned; excess diagnostics add explicit unknown
coverage and cannot turn a failure into `PASS`. Oversized or unsupported input
is rejected, never shortened and labeled complete.

## Evidence boundaries

Generation returns `GENERATED_NOT_VERIFIED`; its semantic/render checks are
`NOT_RUN`. Verification semantic `PASS` proves only the supported supplied
content in the selected exact output under this profile. It does not prove
arbitrary visual equivalence, upstream report quality/completeness, platform
execution identity or authorship. Render all actual synthetic fixture pages
and inspect glyphs, clipping, hierarchy, list markers/order and editability
before claiming visual acceptance. This tool offers no PDF export, table/media
adapter, arbitrary DOCX round trip, deidentifier, rights checker or platform
integration. Caller content remains caller content.
