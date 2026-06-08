def group_paragraphs(layout: list) -> list:
    """
    Convert spans into paragraph objects.
    Each paragraph = {"text", "x_norm", "y_norm", "center"}
    """
    paragraphs = []

    for span in layout:
        text = span.get("text", "").strip()
        if not text:
            continue

        paragraphs.append(
            {
                "text": text,
                "x_norm": span.get("x_norm", 0.0),
                "y_norm": span.get("y_norm", 0.0),
                "doc_y_norm": span.get("doc_y_norm", 0.0),
                "bbox_norm": span.get(
                    "bbox_norm",
                    {"x0": 0.0, "y0": 0.0, "x1": 0.0, "y1": 0.0},
                ),
                "center": span.get("center", (0, 0)),
                "page": span.get("page", 0),
            }
        )

    return paragraphs

# deepseek
