"""CodSpeed benchmarks for docx-scalpel's pure-Python wire (de)serialization.

Every ``DocxSession`` call crosses an NDJSON boundary to the ``docxodus-pyhost``
subprocess. On the Python side, that means two hot, host-independent code
paths run on every single op:

* **Decode** — turning the parsed JSON dict the host returns into the frozen
  dataclasses in ``docx_scalpel.types`` via the ``_from_wire`` classmethods.
  A ``project()`` on a large document decodes an ``anchorIndex`` with thousands
  of entries; a ``grep`` decodes a list of matches each carrying nested run
  fragments; ``docx_diff_get_revisions`` decodes a list of revisions.
* **Encode** — turning option/settings dataclasses into the camelCase wire
  dicts the host reads via the ``to_wire`` methods.

These benchmarks exercise that (de)serialization on realistic, sized payloads.
They need no ``docxodus-pyhost`` binary: the wire dicts are synthesized in
setup (excluded from measurement) so the benchmarks are fully deterministic and
reproducible in CI under CodSpeed's simulation instrument.
"""

from __future__ import annotations

import json
from typing import Any

from docx_scalpel.types import (
    DocxDiffRevision,
    DocxDiffSettings,
    DocxSessionSettings,
    EditResult,
    HtmlOptions,
    MarkdownProjection,
    TextMatch,
    WmlToMarkdownConverterSettings,
)

# ---------------------------------------------------------------------------
# Synthetic wire-payload builders (run in setup — not part of the measurement)
# ---------------------------------------------------------------------------


def _make_anchor_index(n: int) -> dict[str, Any]:
    """An ``anchorIndex`` wire dict with ``n`` block entries, as ``project()``
    returns for a large document."""
    kinds = ("p", "h", "li", "tbl", "tc")
    scopes = ("body", "hdr1", "ftr1", "fn", "cmt")
    index: dict[str, Any] = {}
    for i in range(n):
        kind = kinds[i % len(kinds)]
        scope = scopes[i % len(scopes)]
        unid = f"{i:08x}{'ab3f9c2e' * 3}"
        index[f"{kind}:{scope}:{unid}"] = {
            "kind": kind,
            "scope": scope,
            "unid": unid,
            "partUri": f"/word/{scope}.xml",
            "textPreview": (
                f"Block {i}: the quick brown fox jumps over the lazy dog "
                f"and then continues with more representative body text."
            ),
        }
    return index


def _make_projection_wire(n: int) -> dict[str, Any]:
    lines = [f"## Section {i}\n\nParagraph {i} body text {{#p:body:{i:08x}}}" for i in range(n)]
    return {"markdown": "\n\n".join(lines), "anchorIndex": _make_anchor_index(n)}


def _make_run_fragment(i: int) -> dict[str, Any]:
    return {
        "unid": f"run{i:08x}",
        "text": f"fragment text number {i} with some words",
        "spanInElement": {"start": i * 3, "length": 42},
        "formatting": {
            "bold": i % 2 == 0,
            "italic": i % 3 == 0,
            "underline": False,
            "strike": False,
            "code": i % 5 == 0,
            "color": "FF0000" if i % 4 == 0 else None,
            "hyperlinkUrl": None,
            "runStyle": "Emphasis" if i % 7 == 0 else None,
        },
    }


def _make_text_match(i: int) -> dict[str, Any]:
    return {
        "text": f"match {i} spanning several runs of formatted content",
        "enclosingAnchor": {
            "id": f"p:body:{i:08x}deadbeef",
            "kind": "p",
            "scope": "body",
            "unid": f"{i:08x}deadbeef",
        },
        "span": {"start": i * 5, "length": 51},
        "fragments": [_make_run_fragment(j) for j in range(4)],
        "contextBefore": "... some preceding context text ...",
        "contextAfter": "... some following context text ...",
        "groups": [f"g{i}a", f"g{i}b"],
    }


def _make_matches_wire(n: int) -> list[dict[str, Any]]:
    return [_make_text_match(i) for i in range(n)]


def _make_revision(i: int) -> dict[str, Any]:
    types = ("Inserted", "Deleted", "Moved", "FormatChanged")
    rtype = types[i % len(types)]
    rev: dict[str, Any] = {
        "revisionType": rtype,
        "text": f"revision {i} text content that was changed in the document",
        "author": "Reviewer",
        "date": "2026-07-08T00:00:00Z",
        "leftAnchor": f"p:body:{i:08x}L",
        "rightAnchor": f"p:body:{i:08x}R",
    }
    if rtype == "Moved":
        rev["moveGroupId"] = i
        rev["isMoveSource"] = i % 2 == 0
    if rtype == "FormatChanged":
        rev["formatChange"] = {
            "oldProperties": {"bold": "false", "size": "22"},
            "newProperties": {"bold": "true", "size": "24"},
            "changedPropertyNames": ["bold", "size"],
        }
    return rev


def _make_revisions_wire(n: int) -> list[dict[str, Any]]:
    return [_make_revision(i) for i in range(n)]


def _make_edit_result_wire(n_anchors: int) -> dict[str, Any]:
    def anchors(prefix: str) -> list[dict[str, Any]]:
        return [
            {
                "id": f"p:body:{prefix}{i:08x}",
                "kind": "p",
                "scope": "body",
                "unid": f"{prefix}{i:08x}",
            }
            for i in range(n_anchors)
        ]

    return {
        "success": True,
        "created": anchors("c"),
        "modified": anchors("m"),
        "removed": anchors("r"),
        "patch": {
            "scopeAnchorId": "p:body:scope0001",
            "markdown": "## Patched\n\nRe-projected markdown content." * 8,
        },
    }


# Pre-build payloads once at import (setup cost excluded from measurements).
_PROJECTION_LARGE = _make_projection_wire(2000)
_PROJECTION_JSON = json.dumps(_PROJECTION_LARGE)
_MATCHES = _make_matches_wire(500)
_REVISIONS = _make_revisions_wire(1000)
_EDIT_RESULT = _make_edit_result_wire(200)


# ---------------------------------------------------------------------------
# Decode benchmarks
# ---------------------------------------------------------------------------


def test_decode_large_projection(benchmark) -> None:
    """Decode a 2000-block ``project()`` response into a ``MarkdownProjection``."""
    result = benchmark(MarkdownProjection._from_wire, _PROJECTION_LARGE)
    assert len(result.anchor_index) == 2000


def test_json_parse_and_decode_projection(benchmark) -> None:
    """Full transport decode: ``json.loads`` + ``_from_wire`` on a large projection."""

    def decode() -> MarkdownProjection:
        return MarkdownProjection._from_wire(json.loads(_PROJECTION_JSON))

    result = benchmark(decode)
    assert len(result.anchor_index) == 2000


def test_decode_text_matches(benchmark) -> None:
    """Decode a 500-match ``grep`` response, each with nested run fragments."""

    def decode() -> list[TextMatch]:
        return [TextMatch._from_wire(m) for m in _MATCHES]

    result = benchmark(decode)
    assert len(result) == 500


def test_decode_diff_revisions(benchmark) -> None:
    """Decode a 1000-revision ``docx_diff_get_revisions`` response."""

    def decode() -> list[DocxDiffRevision]:
        return [DocxDiffRevision._from_wire(r) for r in _REVISIONS]

    result = benchmark(decode)
    assert len(result) == 1000


def test_decode_edit_result(benchmark) -> None:
    """Decode a mutation ``EditResult`` envelope with 600 touched anchors."""
    result = benchmark(EditResult._from_wire, _EDIT_RESULT)
    assert result.success


# ---------------------------------------------------------------------------
# Encode benchmarks
# ---------------------------------------------------------------------------


def test_encode_session_settings(benchmark) -> None:
    """Encode ``DocxSessionSettings`` (with nested projection settings) to wire."""
    settings = DocxSessionSettings(
        revision_author="Redline",
        projection_settings=WmlToMarkdownConverterSettings(),
    )
    wire = benchmark(settings.to_wire)
    assert "projectionSettings" in wire


def test_encode_html_options(benchmark) -> None:
    """Encode the full ``HtmlOptions`` bundle to its camelCase wire dict."""
    options = HtmlOptions(
        page_title="Contract",
        comment_render_mode=1,
        pagination_mode=1,
        render_annotations=True,
        document_language="en-US",
    )
    wire = benchmark(options.to_wire)
    assert wire["paginationMode"] == 1


def test_encode_diff_settings(benchmark) -> None:
    """Encode ``DocxDiffSettings`` (sparse non-default emit) to wire."""
    settings = DocxDiffSettings(
        author_for_revisions="Legal",
        case_insensitive=True,
        move_similarity_threshold=0.9,
        move_minimum_word_count=5,
    )
    wire = benchmark(settings.to_wire)
    assert wire["authorForRevisions"] == "Legal"
