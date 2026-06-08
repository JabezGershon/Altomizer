import fitz  # PyMuPDF


def extract_layout(pdf_path: str) -> list:
    """
    Extract text spans from PDF with normalized coordinates.
    Each span = {"text", "x_norm", "y_norm", "center", "page"}
    """
    doc = fitz.open(pdf_path)
    layout = []
    total_pages = max(len(doc), 1)

    for page_num, page in enumerate(doc):
        width = page.rect.width
        height = page.rect.height
        blocks = page.get_text("blocks")

        for b in blocks:
            x0, y0, x1, y1, text, *_ = b
            text = text.strip()

            if not text:
                continue

            layout.append(
                {
                    "text": text,
                    "x_norm": x0 / width,
                    "y_norm": y0 / height,
                    "doc_y_norm": (page_num + (y0 / height)) / total_pages,
                    "bbox_norm": {
                        "x0": x0 / width,
                        "y0": y0 / height,
                        "x1": x1 / width,
                        "y1": y1 / height,
                    },
                    "center": ((x0 + x1) / 2, (y0 + y1) / 2),
                    "page": page_num,
                }
            )

    layout.sort(
        key=lambda item: (
            item.get("page", 0),
            item.get("y_norm", 0),
            item.get("x_norm", 0),
        )
    )
    return layout

# deepseek
