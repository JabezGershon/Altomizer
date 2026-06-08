import re
from difflib import SequenceMatcher


def normalize_text(text: str) -> str:
    text = (text or "").lower().replace("\n", " ")
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"[^a-z0-9=+\-*/^().,:; ]", "", text)
    return text.strip()


def similarity(a: str, b: str) -> float:
    return SequenceMatcher(None, a, b).ratio()


def semantic_similarity(a: str, b: str) -> float:
    return similarity(normalize_text(a), normalize_text(b))


def pair_score(a: dict, b: dict) -> float:
    raw_sim = similarity(a.get("text", ""), b.get("text", ""))
    sem_sim = semantic_similarity(a.get("text", ""), b.get("text", ""))
    doc_y_diff = abs(
        a.get("doc_y_norm", a.get("y_norm", 0))
        - b.get("doc_y_norm", b.get("y_norm", 0))
    )
    local_y_diff = abs(a.get("y_norm", 0) - b.get("y_norm", 0))
    x_diff = abs(a.get("x_norm", 0) - b.get("x_norm", 0))
    page_gap = abs(a.get("page", 0) - b.get("page", 0))
    role_a = a.get("role")
    role_b = b.get("role")
    reading_a = a.get("reading_order")
    reading_b = b.get("reading_order")
    sequence_a = a.get("sequence_order")
    sequence_b = b.get("sequence_order")
    reading_gap = None
    if isinstance(reading_a, int) and isinstance(reading_b, int):
        reading_gap = abs(reading_a - reading_b)
    sequence_gap = None
    if isinstance(sequence_a, int) and isinstance(sequence_b, int):
        sequence_gap = abs(sequence_a - sequence_b)

    # Strongly favor textual/semantic agreement, but keep matches local in reading order.
    score = (
        (raw_sim * 0.34)
        + (sem_sim * 0.46)
        + ((1 - min(local_y_diff, 1)) * 0.12)
        + ((1 - min(x_diff, 1)) * 0.05)
        + ((1 - min(doc_y_diff, 1)) * 0.02)
        + ((1 / (1 + page_gap)) * 0.01)
    )
    if role_a and role_b:
        if role_a == role_b:
            score += 0.05
        elif {role_a, role_b} & {"image", "equation"} and role_a != role_b:
            score -= 0.08
    if reading_gap is not None:
        score += 0.05 / (1 + reading_gap)
    if sequence_gap is not None:
        score += 0.08 / (1 + sequence_gap)
    return round(score, 6)


def gap_penalty(item: dict) -> float:
    text = normalize_text(item.get("text", ""))
    if not text:
        return 0.08
    if len(text) <= 6:
        return 0.12
    return 0.18


def sort_elements(elements: list) -> list:
    return sorted(
        elements,
        key=lambda item: (
            item.get("sequence_order", 10**9),
            item.get("page", 0),
            item.get("reading_order", 10**9),
            item.get("y_norm", 0),
            item.get("x_norm", 0),
        ),
    )


def match_elements(docA, docB):
    """
    Order-preserving alignment for repeated labels/equations.

    Greedy matching caused duplicate-looking formulas to attach to the wrong
    occurrence. This dynamic-programming alignment preserves reading order and
    allows skips with a penalty instead of crossing matches.
    """
    left = sort_elements(docA)
    right = sort_elements(docB)

    m = len(left)
    n = len(right)
    if m == 0 or n == 0:
        return []

    dp = [[0.0] * (n + 1) for _ in range(m + 1)]
    move = [[None] * (n + 1) for _ in range(m + 1)]

    for i in range(1, m + 1):
        dp[i][0] = dp[i - 1][0] - gap_penalty(left[i - 1])
        move[i][0] = "up"

    for j in range(1, n + 1):
        dp[0][j] = dp[0][j - 1] - gap_penalty(right[j - 1])
        move[0][j] = "left"

    for i in range(1, m + 1):
        for j in range(1, n + 1):
            pairing = dp[i - 1][j - 1] + pair_score(left[i - 1], right[j - 1])
            skip_left = dp[i - 1][j] - gap_penalty(left[i - 1])
            skip_right = dp[i][j - 1] - gap_penalty(right[j - 1])

            best = pairing
            best_move = "diag"

            if skip_left > best:
                best = skip_left
                best_move = "up"

            if skip_right > best:
                best = skip_right
                best_move = "left"

            dp[i][j] = best
            move[i][j] = best_move

    pairs = []
    i = m
    j = n

    while i > 0 and j > 0:
        action = move[i][j]

        if action == "diag":
            a = left[i - 1]
            b = right[j - 1]
            score = pair_score(a, b)

            # Keep only plausible alignments; low-confidence pairs are better skipped.
            if score >= 0.45:
                pairs.append((a, b, score))
            i -= 1
            j -= 1
        elif action == "up":
            i -= 1
        else:
            j -= 1

    pairs.reverse()
    return pairs
