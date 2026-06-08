import re
from difflib import SequenceMatcher


def normalize_text(text: str) -> str:
    text = (text or "").lower()
    text = text.replace("\n", " ")
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"[_`~|]", " ", text)
    text = re.sub(r"[^a-z0-9=+\-*/^().,:; ]", "", text)
    return text.strip()


def semantic_similarity(a: str, b: str) -> float:
    return SequenceMatcher(None, normalize_text(a), normalize_text(b)).ratio()


def is_equation_like(text: str) -> bool:
    normalized = normalize_text(text)
    symbols = ["=", "+", "-", "^", "/", "(", ")", "x", "y", "a2", "b2", "c2"]
    score = sum(symbol in normalized for symbol in symbols)
    digit_count = sum(char.isdigit() for char in normalized)
    return score >= 3 or (score >= 2 and digit_count >= 2)


def build_candidate_result(idx: int, a: dict, b: dict, sim: float) -> dict | None:
    accepted_similarity = 0.85
    semantic_sim = semantic_similarity(a.get("text", ""), b.get("text", ""))
    pos_diff = abs(a.get("y_norm", 0) - b.get("y_norm", 0))
    indent_diff = abs(a.get("x_norm", 0) - b.get("x_norm", 0))
    page_gap = abs(a.get("page", 0) - b.get("page", 0))
    equation_like = is_equation_like(a.get("text", "")) or is_equation_like(b.get("text", ""))
    semantically_same = semantic_sim >= 0.93

    issues = []
    classification = "content-change"

    if page_gap >= 1:
        issues.append("CROSS_PAGE_ALIGNMENT")

    if sim < accepted_similarity and semantic_sim < 0.78 and not equation_like:
        issues.append("TEXT_MISMATCH")
    elif semantic_sim < 0.6 and not equation_like:
        issues.append("CONTENT_CHANGE")

    if indent_diff > 0.28 and not semantically_same:
        issues.append("INDENTATION_MISMATCH")

    if pos_diff > 0.30:
        issues.append("MAJOR_POSITION_SHIFT")
    elif pos_diff > 0.15:
        issues.append("PAGE_SHIFT")
    elif pos_diff > 0.07:
        issues.append("MINOR_POSITION_SHIFT")

    if equation_like and semantic_sim < 0.7:
        issues.append("EQUATION_OR_GRAPHIC_CHANGE")
        classification = "visual-structure-change"

    if semantically_same and any(
        issue in issues
        for issue in {"MAJOR_POSITION_SHIFT", "PAGE_SHIFT", "MINOR_POSITION_SHIFT"}
    ):
        classification = "reflow-candidate"

    if not issues:
        return None

    return {
        "index": idx,
        "text": a.get("text", "")[:90],
        "issues": issues,
        "distance": round(pos_diff, 4),
        "similarity": round(sim, 4),
        "semantic_similarity": round(semantic_sim, 4),
        "indent_delta": round(indent_diff, 4),
        "page_gap": page_gap,
        "classification": classification,
        "left": {
            "text": a.get("text", ""),
            "page": a.get("page", 0),
            "bbox_norm": a.get("bbox_norm", {}),
            "x_norm": a.get("x_norm", 0),
            "y_norm": a.get("y_norm", 0),
            "doc_y_norm": a.get("doc_y_norm", 0),
        },
        "right": {
            "text": b.get("text", ""),
            "page": b.get("page", 0),
            "bbox_norm": b.get("bbox_norm", {}),
            "x_norm": b.get("x_norm", 0),
            "y_norm": b.get("y_norm", 0),
            "doc_y_norm": b.get("doc_y_norm", 0),
        },
    }


def collapse_reflow_candidates(candidates: list) -> list:
    final_results = []
    idx = 0

    while idx < len(candidates):
        current = candidates[idx]

        if current.get("classification") != "reflow-candidate":
            final_results.append(current)
            idx += 1
            continue

        cluster = [current]
        shift_direction = 1 if current.get("left", {}).get("y_norm", 0) <= current.get("right", {}).get("y_norm", 0) else -1
        next_idx = idx + 1

        while next_idx < len(candidates):
            candidate = candidates[next_idx]
            same_page = (
                candidate.get("classification") == "reflow-candidate"
                and candidate.get("left", {}).get("page") == current.get("left", {}).get("page")
                and candidate.get("right", {}).get("page") == current.get("right", {}).get("page")
            )
            candidate_direction = 1 if candidate.get("left", {}).get("y_norm", 0) <= candidate.get("right", {}).get("y_norm", 0) else -1

            if not same_page or candidate_direction != shift_direction:
                break

            cluster.append(candidate)
            next_idx += 1

        if len(cluster) >= 2:
            lead = cluster[0]
            lead["issues"] = ["REFLOW_AFTER_VISUAL_CHANGE"]
            lead["classification"] = "reflow"
            lead["reflow_span"] = len(cluster)
            lead["distance"] = round(max(item.get("distance", 0) for item in cluster), 4)
            final_results.append(lead)
        else:
            isolated = cluster[0]
            isolated["issues"] = ["LAYOUT_SHIFT"]
            isolated["classification"] = "layout-shift"
            final_results.append(isolated)

        idx = next_idx

    return final_results


def compare_pairs(pairs: list) -> list:
    candidates = []

    for idx, (a, b, sim) in enumerate(pairs):
        candidate = build_candidate_result(idx, a, b, sim)
        if candidate is not None:
            candidates.append(candidate)

    return collapse_reflow_candidates(candidates)
