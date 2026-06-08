from __future__ import annotations


class DocumentPreserver:
    """Compatibility wrapper for flows that expect a snapshot/apply interface.

    The list-correction processor only edits paragraph/list metadata, and
    `python-docx` already preserves untouched package parts on save for this
    workflow. This shim keeps the original calling convention without adding a
    hard dependency on an external `preserver` package.
    """

    def __init__(self, _document) -> None:
        self._document = _document

    def snapshot(self, include_hf: bool = True) -> None:
        _ = include_hf

    def apply(self) -> None:
        return None
