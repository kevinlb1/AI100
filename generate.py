from __future__ import annotations

import random
from typing import List, Tuple


def apply_fixed_vetoes_least_preferred(
    v: List[List[int]],
    seed: int,
    veto_count: int | None = None,
    enforce_self_five: bool = True,
) -> None:
    """Mutate v in-place to impose a fixed number of vetoes per student."""
    n = len(v)
    if any(len(row) != n for row in v):
        raise ValueError("v must be an n x n matrix")

    if veto_count is None:
        veto_count = n // 4
    if veto_count < 0:
        raise ValueError("veto_count must be non-negative")
    if n >= 2 and veto_count > n - 2:
        veto_count = n - 2

    for i in range(n):
        rng = random.Random((seed + 1) * 1_000_003 + (i + 11) * 97_331)
        candidates = [j for j in range(n) if j != i]
        rng.shuffle(candidates)
        candidates.sort(key=lambda j: v[i][j])
        for j in candidates[:veto_count]:
            v[i][j] = 0
        if enforce_self_five:
            v[i][i] = 5


def generate_preferences_random(
    n: int,
    C: int,
    seed: int,
) -> Tuple[List[List[int]], List[int], List[List[int]], List[int]]:
    if n <= 0:
        raise ValueError("n must be positive.")
    if C <= 0:
        C = 1

    rng = random.Random(seed)
    topics = range(n)
    v: List[List[int]] = []
    for i in range(n):
        row = [rng.randint(1, 5) for _ in topics]
        row[i] = 5
        v.append(row)

    apply_fixed_vetoes_least_preferred(v=v, seed=seed, veto_count=n // 4, enforce_self_five=True)
    r = list(range(n))
    cat = [j % C for j in topics]
    base: List[List[int]] = []
    return v, r, base, cat


def generate_preferences_by_category(
    n: int,
    C: int,
    seed: int,
) -> Tuple[List[List[int]], List[int], List[List[int]], List[int]]:
    if n <= 0:
        raise ValueError("n must be positive")
    if C <= 0:
        raise ValueError("C must be positive.")

    rng = random.Random(seed)
    topics = range(n)
    r = list(range(n))

    base: List[List[int]] = []
    fav_cat: List[int] = []
    for _ in range(n):
        fav = rng.randrange(C)
        fav_cat.append(fav)
        row = [5 if c == fav else rng.randint(1, 4) for c in range(C)]
        base.append(row)

    cat = [fav_cat[j] for j in topics]
    v: List[List[int]] = []
    for i in range(n):
        row: List[int] = []
        for j in topics:
            noise = rng.choice([-1, 0, 1])
            score = base[i][cat[j]] + noise
            row.append(max(1, min(5, score)))
        row[i] = 5
        v.append(row)

    apply_fixed_vetoes_least_preferred(v=v, seed=seed, veto_count=n // 4, enforce_self_five=True)
    return v, r, base, cat


def generate_preferences_by_category_mode3(
    n: int,
    C: int,
    seed: int,
) -> Tuple[List[List[int]], List[int], List[List[int]], List[int]]:
    if n <= 0:
        raise ValueError("n must be positive")
    if C <= 0:
        raise ValueError("C must be positive.")

    rng = random.Random(seed)
    topics = range(n)
    r = list(range(n))

    def weighted_choice(values: List[int], weights: List[int]) -> int:
        total = sum(weights)
        x = rng.randrange(total)
        acc = 0
        for value, weight in zip(values, weights):
            acc += weight
            if x < acc:
                return value
        return values[-1]

    pref_vals = [1, 2, 3, 4, 5]
    pref_wts = [1, 4, 10, 4, 1]
    base: List[List[int]] = [[weighted_choice(pref_vals, pref_wts) for _ in range(C)] for _ in range(n)]

    cat: List[int] = []
    for j in topics:
        best = max(base[j])
        best_cs = [c for c in range(C) if base[j][c] == best]
        cat.append(rng.choice(best_cs))

    noise_vals = [-1, 0, 1]
    noise_wts = [1, 4, 1]
    v: List[List[int]] = []
    for i in range(n):
        row: List[int] = []
        for j in topics:
            noise = weighted_choice(noise_vals, noise_wts)
            score = base[i][cat[j]] + noise
            row.append(max(1, min(5, score)))
        row[i] = 5
        v.append(row)

    apply_fixed_vetoes_least_preferred(v=v, seed=seed, veto_count=n // 4, enforce_self_five=True)
    return v, r, base, cat


def generate_preferences_by_category_uniform_real_binned(
    n: int,
    C: int,
    seed: int,
) -> Tuple[List[List[int]], List[int], List[List[float]], List[int]]:
    if n <= 0:
        raise ValueError("n must be positive")
    if C <= 0:
        raise ValueError("C must be positive.")

    rng = random.Random(seed)
    topics = range(n)
    r = list(range(n))

    base_real: List[List[float]] = [[rng.uniform(1.0, 5.0) for _ in range(C)] for _ in range(n)]
    cat: List[int] = []
    for j in topics:
        best = max(base_real[j])
        best_cs = [c for c in range(C) if abs(base_real[j][c] - best) <= 1e-12]
        cat.append(rng.choice(best_cs))

    real_scores: List[List[float]] = []
    for i in range(n):
        own_cat = cat[i]
        base_i = list(base_real[i])
        base_i[own_cat] = 4.5
        row: List[float] = []
        for j in topics:
            row.append(base_i[cat[j]] + rng.uniform(-1.0, 1.0))
        real_scores.append(row)

    v: List[List[int]] = []
    veto_count = n // 4
    for i in range(n):
        order_asc = sorted(topics, key=lambda j: (real_scores[i][j], rng.random()))
        veto = set(order_asc[:veto_count])
        remaining = [j for j in topics if j not in veto]
        order_desc = sorted(remaining, key=lambda j: (real_scores[i][j], rng.random()), reverse=True)
        m = len(order_desc)
        row = [0] * n
        if m > 0:
            for rank_idx, j in enumerate(order_desc):
                bucket = int((rank_idx * 5) / m)
                row[j] = 5 - bucket
        v.append(row)

    return v, r, base_real, cat
