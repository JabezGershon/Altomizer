def is_valid_paragraph(p: dict) -> bool:
    """
    Filter out junk paragraphs (too short, missing layout, headers/footers).
    """
    if not isinstance(p, dict):
        return False

    text = p.get("text", "").strip()
    if len(text) < 5:
        return False

    # Safely get normalized Y (if missing, assume valid)
    y = p.get("y_norm")
    if y is not None and (y < 0.05 or y > 0.95):
        return False

    return True

# deepseek