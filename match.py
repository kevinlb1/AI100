from __future__ import annotations

from ortools.sat.python import cp_model
import random
import math
import csv
from statistics import median
try:
    import numpy as np  # type: ignore
except Exception:  # pragma: no cover
    np = None
try:
    import matplotlib.pyplot as plt  # type: ignore
except Exception:  # pragma: no cover
    plt = None
import multiprocessing
import time
from typing import List, Tuple, Dict, Any


def choose_even_k(n: int, M: int) -> Tuple[int, bool]:
    """Choose the smallest even K such that K(M-1) <= n <= K*M.

    Returns (K, feasible_flag). If no such even K exists, returns the smallest
    even >= ceil(n/M) and feasible_flag=False.
    """
    if M <= 1:
        raise ValueError("M must be >= 2.")
    k_min = math.ceil(n / M)
    k_max = n // (M - 1)  # floor(n/(M-1))

    k = k_min if k_min % 2 == 0 else k_min + 1
    if k <= k_max:
        return k, True
    return k, False


def visualize_results(student_ids, main_scores, shadow_scores, n: int):
    """Simple satisfaction chart."""
    if plt is None:
        return
    x = list(range(len(student_ids)))
    width = 0.35

    fig, ax = plt.subplots(figsize=(14, 6))
    ax.bar([xi - width / 2 for xi in x], main_scores, width, label='Main', color='#27ae60')
    ax.bar([xi + width / 2 for xi in x], shadow_scores, width, label='Shadow', color='#2980b9')

    ax.set_ylabel('Preference Score')
    ax.set_title(f'Student Satisfaction (n={n})')
    ax.set_xticks(x)
    ax.set_xticklabels(student_ids)
    ax.legend()
    ax.set_ylim(0, 6)

    plt.tight_layout()
    plt.show()


def apply_fixed_vetoes_least_preferred(
    v: List[List[int]],
    seed: int,
    veto_count: int | None = None,
    enforce_self_five: bool = True,
) -> None:
    """Mutate v in-place to impose a fixed number of vetoes per student.

    Procedure (per student i):
      1) Start from fully-scored preferences v[i][j] in {1..5}.
      2) Zero out exactly `veto_count` topics among j != i with smallest score.
      3) Break ties randomly (reproducibly from `seed`).

    By default veto_count = n//4.

    Notes:
      - We never veto a student's own topic i.
      - If enforce_self_five is True, we set v[i][i]=5 after vetoing as well.
    """
    n = len(v)
    if any(len(row) != n for row in v):
        raise ValueError("v must be an n x n matrix")

    if veto_count is None:
        veto_count = n // 4
    if veto_count < 0:
        raise ValueError("veto_count must be non-negative")
    if n >= 2 and veto_count > n - 2:
        # With main != shadow, each student should generally have >=2 allowed topics.
        veto_count = n - 2

    for i in range(n):
        # Deterministic per-student RNG for tie-breaking.
        rng = random.Random((seed + 1) * 1_000_003 + (i + 11) * 97_331)

        candidates = [j for j in range(n) if j != i]
        rng.shuffle(candidates)  # random tie-breaking via stable sort
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
    """Legacy-ish generator: random scores then fixed-count vetoes.

    For each student i:
      - Draw raw v[i][j] ~ Uniform{1..5} independently for each topic j
      - Force v[i][i] = 5
      - Impose exactly n//4 vetoes by zeroing out the least-preferred topics
        (ties broken randomly)

    Categories are produced only for reporting; they do not affect preferences in this mode.

    Returns: (v, r, base, cat)
      - v[i][j] in {0..5}
      - r[i] = i
      - base is empty (not used)
      - cat[j] in {0..C-1} (topic categories for reporting)
    """
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
    """Category-based generator (C categories) + fixed-count vetoes.

    - Each student i has base preference for each category c: base[i][c] in {1..5}.
      We generate this by giving each student a favorite category at 5, and others
      in {1..4}.
    - Each topic j has a category cat[j]. Since topics correspond to proposers in this
      harness (r_j = j), we set cat[j] to proposer j's favorite category.
    - Raw preference for topic j is base[i][cat[j]] + noise, where noise ∈ {-1,0,1}.
      Scores are clipped to {1..5}.
    - After generating all raw scores, each student vetoes exactly n//4 topics by
      zeroing out the least-preferred topics (ties broken randomly).

    Returns: (v, r, base, cat)
    """
    if n <= 0:
        raise ValueError("n must be positive")
    if C <= 0:
        raise ValueError("C must be positive.")

    rng = random.Random(seed)
    topics = range(n)

    r = list(range(n))

    # Student category preferences
    base: List[List[int]] = []
    fav_cat: List[int] = []
    for _i in range(n):
        fav = rng.randrange(C)
        fav_cat.append(fav)
        row = []
        for c in range(C):
            if c == fav:
                row.append(5)
            else:
                row.append(rng.randint(1, 4))
        base.append(row)

    # Topic categories: topic j proposed by student j
    cat = [fav_cat[j] for j in topics]

    # Raw topic-level preferences with +/-1 noise
    v: List[List[int]] = []
    for i in range(n):
        row: List[int] = []
        for j in topics:
            noise = rng.choice([-1, 0, 1])
            score = base[i][cat[j]] + noise
            score = max(1, min(5, score))
            row.append(score)
        row[i] = 5
        v.append(row)

    apply_fixed_vetoes_least_preferred(v=v, seed=seed, veto_count=n // 4, enforce_self_five=True)

    return v, r, base, cat

def generate_preferences_by_category_mode3(
    n: int,
    C: int,
    seed: int,
) -> Tuple[List[List[int]], List[int], List[List[int]], List[int]]:
    """Category-based generator with scores peaked at 3 + fixed-count vetoes.

    This is meant to be a more "realistic" harness distribution:
      - Students have base preferences over categories, but those preferences are
        centered around neutral (3). Extreme likes/dislikes (5/1) are uncommon.
      - Topic-level scores are base[i][cat[j]] plus small noise in {-1,0,1}, with 0
        more likely than ±1 (so scores stay near the base).

    Concretely:
      - base[i][c] is drawn from a discrete distribution on {1..5} with mode at 3
        (weights 1:1, 2:4, 3:10, 4:4, 5:1).
      - noise is drawn from {-1,0,1} with mode at 0 (weights -1:1, 0:4, +1:1).
      - Each topic j has category cat[j] equal to the (tie-broken) argmax of the
        proposer j's base preferences (since r_j = j in this harness).
      - After generating all raw scores, each student vetoes exactly n//4 topics by
        zeroing out the least-preferred topics (ties broken randomly).

    Returns: (v, r, base, cat)
    """
    if n <= 0:
        raise ValueError("n must be positive")
    if C <= 0:
        raise ValueError("C must be positive.")

    rng = random.Random(seed)
    topics = range(n)
    r = list(range(n))

    # Helper: sample from a small discrete distribution specified by weights.
    def weighted_choice(values: List[int], weights: List[int]) -> int:
        total = sum(weights)
        x = rng.randrange(total)
        acc = 0
        for v_, w_ in zip(values, weights):
            acc += w_
            if x < acc:
                return v_
        return values[-1]

    # Student category preferences, peaked at 3.
    pref_vals = [1, 2, 3, 4, 5]
    pref_wts = [1, 4, 10, 4, 1]
    base: List[List[int]] = []
    for _i in range(n):
        base.append([weighted_choice(pref_vals, pref_wts) for _c in range(C)])

    # Topic categories: topic j proposed by student j. Assign category as argmax base[j][c].
    cat: List[int] = []
    for j in topics:
        best = max(base[j])
        best_cs = [c for c in range(C) if base[j][c] == best]
        cat.append(rng.choice(best_cs))

    # Raw topic-level preferences with small noise, mostly 0.
    noise_vals = [-1, 0, 1]
    noise_wts = [1, 4, 1]
    v: List[List[int]] = []
    for i in range(n):
        row: List[int] = []
        for j in topics:
            noise = weighted_choice(noise_vals, noise_wts)
            score = base[i][cat[j]] + noise
            score = max(1, min(5, score))
            row.append(score)
        row[i] = 5
        v.append(row)

    apply_fixed_vetoes_least_preferred(v=v, seed=seed, veto_count=n // 4, enforce_self_five=True)
    return v, r, base, cat



def generate_preferences_by_category_uniform_real_binned(
    n: int,
    C: int,
    seed: int,
) -> Tuple[List[List[int]], List[int], List[List[float]], List[int]]:
    """Category-based generator with real-valued utilities, vetoes, and quintile binning.

    Procedure:
      - Draw base_real[i][c] ~ Uniform(1,5) for each student/category.
      - Topic category cat[j] is proposer j's argmax category (tie-broken randomly).
      - For each student i, set base for own proposal category cat[i] to 4.5.
      - Topic real preference: base_i[cat[j]] + Uniform(-1,1).
      - Veto least favorite quarter of topics (n//4): score 0.
      - Remaining topics are ranked by real preference and binned evenly into
        five groups: top fifth->5, next->4, ..., last->1.
    """
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
                bucket = int((rank_idx * 5) / m)  # 0..4
                row[j] = 5 - bucket
        v.append(row)

    return v, r, base_real, cat


class ProgressSummary(cp_model.CpSolverSolutionCallback):
    """Print a compact progress line each time CP-SAT finds a new incumbent."""

    def __init__(
        self,
        t0: float,
        utility10_var: cp_model.IntVar,
        penalty_var: cp_model.IntVar,
        maximize: bool = True,
        print_header: bool = True,
    ):
        super().__init__()
        self._t0 = t0
        self._utility10_var = utility10_var
        self._penalty_var = penalty_var
        self._maximize = maximize
        self._print_header = print_header
        self._last_obj = None
        self._sols = 0

    def OnSolutionCallback(self) -> None:  # noqa: N802
        if self._print_header:
            print("\nProgress (new incumbents):")
            print("  time(s)   util   pen        obj      bound      gap   sols")
            self._print_header = False

        elapsed = time.perf_counter() - self._t0
        util10 = int(self.Value(self._utility10_var))
        pen = int(self.Value(self._penalty_var))
        # ObjectiveValue() naming differs across OR-Tools versions.
        try:
            obj = float(self.ObjectiveValue())
        except Exception:
            try:
                obj = float(self.objective_value())
            except Exception:
                obj = float("nan")

        # BestObjectiveBound() exists on modern OR-Tools; guard just in case.
        try:
            bound = float(self.BestObjectiveBound())
        except Exception:
            try:
                bound = float(self.best_objective_bound())
            except Exception:
                bound = float('nan')

        if self._maximize:
            gap = bound - obj
        else:
            gap = obj - bound

        self._sols += 1

        sols = self._sols

        # In optimization, OR-Tools typically calls this only on improving incumbents,
        # but guard anyway.
        if self._last_obj is None or (self._maximize and obj > self._last_obj + 1e-9) or ((not self._maximize) and obj < self._last_obj - 1e-9):
            self._last_obj = obj
            util = util10 / 10.0
            # Bound/obj are integer-valued in our models; print as ints when finite.
            def fmt_num(x: float) -> str:
                return "   n/a" if math.isnan(x) else f"{int(round(x)):9d}"

            bound_s = "   n/a" if math.isnan(bound) else f"{int(round(bound)):9d}"
            gap_s = "   n/a" if math.isnan(gap) else f"{int(round(gap)):7d}"

            print(f"{elapsed:8.1f} {util:6.1f} {pen:5d} {int(round(obj)):10d} {bound_s} {gap_s} {sols:6d}")

    # Compatibility: some OR-Tools examples (and older wrappers) exposed NumSolutions()
    # on the callback; newer wrappers do not. We track solution count ourselves.
    def NumSolutions(self) -> int:  # noqa: N802
        return int(self._sols)

    @property
    def sols(self) -> int:
        return int(self._sols)


def solve_lab_ortools(
    n: int = 40,
    M: int = 4,
    C: int = 8,
    filename: str = "assignments_ortools.csv",
    time_limit_s: float | None = None,
    seed: int = 42,
    pref_mode: str = "category_uniform",  # "category_uniform", "category", "category_mode3", or "random"
    lexicographic_overlap_tiebreak: bool = True,
    weight_W: int = 1000,
    plot: bool = True,
    progress_summary: bool = True,
    raw_cpsat_log: bool = False,
    workers: int | None = None,
):
    """CP-SAT harness.

    - Uses an O(n^2) overlap encoding.
    - Supports four preference generators (switch with pref_mode):
        * pref_mode="random": random raw topic utilities (1..5) + fixed-count vetoes
        * pref_mode="category": category utilities with ±1 noise + fixed-count vetoes
        * pref_mode="category_mode3": category utilities peaked at 3 + fixed-count vetoes
        * pref_mode="category_uniform": real-valued category utilities,
          topic modifiers ±1, veto least quarter, then bin remaining topics into
          quintiles mapped to scores 5..1

    Veto policy (both modes): after generating all raw scores, each student vetoes
    exactly n//4 topics by zeroing out the least-preferred topics, breaking ties randomly.

    Objective scaling:
      - utility10 = sum_i (20 * MainScore_i + 10 * ShadowScore_i)   (i.e., 10× LaTeX scale)
      - penalty  = sum_{i<k} z_{i,k}
      - If lexicographic_overlap_tiebreak=True, maximize W*utility10 - penalty.
        A sufficient bound for strict tie-breaking is W > max_penalty, where
          max_penalty <= n*(M-1)/2.

    Progress reporting:
      - If progress_summary=True, prints a compact line each time a new incumbent is found.
      - If raw_cpsat_log=True, also enables the built-in CP-SAT search log.
    """
    if M < 2:
        raise ValueError("M must be >= 2")

    cpu_threads = multiprocessing.cpu_count()
    workers_used = cpu_threads if workers is None else int(workers)
    if workers_used <= 0:
        raise ValueError("workers must be a positive integer")
    print("--- Group Assignment Harness (Google OR-Tools CP-SAT) ---")
    print(f"Hardware: {cpu_threads} logical cores; using {workers_used} workers")
    print(f"Config: n={n}, M={M}, C={C}, pref_mode={pref_mode}, seed={seed}")
    print(f"Time limit: {time_limit_s if time_limit_s is not None else 'none'}")
    print(f"Progress summary: {'on' if progress_summary else 'off'}")
    print(f"Raw CP-SAT log: {'on' if raw_cpsat_log else 'off'}")
    print(f"Veto policy: exactly n//4 = {n//4} zeros per student (least preferred; ties random)")

    K, k_feasible = choose_even_k(n, M)
    raw_k = math.ceil(n / M)
    print(f"K selection: ceil(n/M)={raw_k}, chosen K={K} (even).")
    if not k_feasible:
        print(
            f"WARNING: No even K satisfies K(M-1) <= n <= KM for n={n}, M={M}. "
            f"With the strict min-size M-1 constraint, the model may be infeasible."
        )

    students = range(n)
    topics = range(n)

    # --- Data generation (harness) ---
    if pref_mode == "category":
        v, r, base, cat = generate_preferences_by_category(n=n, C=C, seed=seed)
    elif pref_mode == "category_mode3":
        v, r, base, cat = generate_preferences_by_category_mode3(n=n, C=C, seed=seed)
    elif pref_mode == "random":
        v, r, base, cat = generate_preferences_random(n=n, C=C, seed=seed)
    elif pref_mode in {"category_uniform", "category_uniform_real_binned"}:
        v, r, base, cat = generate_preferences_by_category_uniform_real_binned(n=n, C=C, seed=seed)
    else:
        raise ValueError('pref_mode must be one of: "category_uniform", "category", "category_mode3", "random"')

    # Quick stats
    zeros_per_student = [sum(1 for j in topics if v[i][j] == 0) for i in students]
    print(
        f"Veto stats: min={min(zeros_per_student)}, median={int(median(zeros_per_student))}, max={max(zeros_per_student)} zeros per student"
    )

    model = cp_model.CpModel()

    # --- Topic-level variables ---
    g = [model.NewBoolVar(f"g_{j}") for j in topics]  # active topics
    p = [model.NewBoolVar(f"p_{j}") for j in topics]  # partition: 1=A, 0=B

    # Symmetry breaking: fix one partition bit to remove A/B flip symmetry.
    model.Add(p[1] == 1) if n >= 2 else model.Add(p[0] == 1)

    # Preprocessing cut: if too few students are eligible for a topic, it cannot be active.
    eligible = [sum(1 for i in students if v[i][j] > 0) for j in topics]
    for j in topics:
        if eligible[j] < (M - 1):
            model.Add(g[j] == 0)

    model.Add(sum(g) == K)

    # Partition balance: sum_j (g_j AND p_j) == K/2
    u = []
    for j in topics:
        uj = model.NewBoolVar(f"u_{j}")
        model.Add(uj <= g[j])
        model.Add(uj <= p[j])
        model.Add(uj >= g[j] + p[j] - 1)
        u.append(uj)
    model.Add(sum(u) == K // 2)

    # --- Student-level variables ---
    Main = []
    Shadow = []
    MainScore = []
    ShadowScore = []

    # Reified equalities for capacity counting
    main_is: Dict[Tuple[int, int], cp_model.IntVar] = {}
    shadow_is: Dict[Tuple[int, int], cp_model.IntVar] = {}

    for i in students:
        allowed = [j for j in topics if v[i][j] > 0]
        if len(allowed) < 2:
            raise ValueError(
                f"Student {i} has fewer than 2 non-veto topics (needs main != shadow)."
            )

        main_i = model.NewIntVarFromDomain(cp_model.Domain.FromValues(allowed), f"Main_{i}")
        shadow_i = model.NewIntVarFromDomain(cp_model.Domain.FromValues(allowed), f"Shadow_{i}")
        Main.append(main_i)
        Shadow.append(shadow_i)

        model.Add(main_i != shadow_i)

        # Topic must be active: g[Main_i] = 1, g[Shadow_i] = 1
        g_main_i = model.NewIntVar(0, 1, f"gMain_{i}")
        g_shadow_i = model.NewIntVar(0, 1, f"gShadow_{i}")
        model.AddElement(main_i, g, g_main_i)
        model.AddElement(shadow_i, g, g_shadow_i)
        model.Add(g_main_i == 1)
        model.Add(g_shadow_i == 1)

        # Opposite partition requirement: p[Main_i] + p[Shadow_i] == 1
        p_main_i = model.NewIntVar(0, 1, f"pMain_{i}")
        p_shadow_i = model.NewIntVar(0, 1, f"pShadow_{i}")
        model.AddElement(main_i, p, p_main_i)
        model.AddElement(shadow_i, p, p_shadow_i)
        model.Add(p_main_i + p_shadow_i == 1)

        # Proposer guarantee: if topic r_i is active, then student i's main topic is r_i.
        model.Add(main_i == r[i]).OnlyEnforceIf(g[r[i]])

        # Preference scores via element
        ms = model.NewIntVar(0, 5, f"MainScore_{i}")
        ss = model.NewIntVar(0, 5, f"ShadowScore_{i}")
        model.AddElement(main_i, v[i], ms)
        model.AddElement(shadow_i, v[i], ss)
        MainScore.append(ms)
        ShadowScore.append(ss)

        # Reified equalities for counting (only for allowed j)
        for j in allowed:
            b = model.NewBoolVar(f"main_is_{i}_{j}")
            model.Add(main_i == j).OnlyEnforceIf(b)
            model.Add(main_i != j).OnlyEnforceIf(b.Not())
            main_is[(i, j)] = b

            cvar = model.NewBoolVar(f"shadow_is_{i}_{j}")
            model.Add(shadow_i == j).OnlyEnforceIf(cvar)
            model.Add(shadow_i != j).OnlyEnforceIf(cvar.Not())
            shadow_is[(i, j)] = cvar

    # --- Capacity constraints ---
    for j in topics:
        main_count = sum(main_is[(i, j)] for i in students if (i, j) in main_is)
        shadow_count = sum(shadow_is[(i, j)] for i in students if (i, j) in shadow_is)
        model.Add(main_count <= M * g[j])
        model.Add(main_count >= (M - 1) * g[j])
        model.Add(shadow_count <= M * g[j])
        model.Add(shadow_count >= (M - 1) * g[j])

    # --- Overlap penalty in O(n^2) ---
    overlap = []
    for i in students:
        for k in range(i + 1, n):
            eq_main = model.NewBoolVar(f"eqMain_{i}_{k}")
            eq_shadow = model.NewBoolVar(f"eqShadow_{i}_{k}")

            model.Add(Main[i] == Main[k]).OnlyEnforceIf(eq_main)
            model.Add(Main[i] != Main[k]).OnlyEnforceIf(eq_main.Not())

            model.Add(Shadow[i] == Shadow[k]).OnlyEnforceIf(eq_shadow)
            model.Add(Shadow[i] != Shadow[k]).OnlyEnforceIf(eq_shadow.Not())

            z = model.NewBoolVar(f"z_{i}_{k}")
            model.Add(z <= eq_main)
            model.Add(z <= eq_shadow)
            model.Add(z >= eq_main + eq_shadow - 1)
            overlap.append(z)

    # --- Objective vars (for progress reporting) ---
    # utility10 is at most (20+10)*5*n = 150n
    utility10_ub = 150 * n
    utility10_var = model.NewIntVar(0, utility10_ub, "utility10")
    model.Add(utility10_var == sum(20 * MainScore[i] + 10 * ShadowScore[i] for i in students))

    # penalty is at most n*(n-1)/2, but typically much smaller.
    penalty_ub = n * (n - 1) // 2
    penalty_var = model.NewIntVar(0, penalty_ub, "penalty")
    model.Add(penalty_var == sum(overlap))

    W_used = None
    if lexicographic_overlap_tiebreak:
        min_safe_W = (n * (M - 1)) // 2 + 1
        W_used = max(weight_W, min_safe_W)
        if W_used != weight_W:
            print(f"WARNING: weight_W={weight_W} too small for strict tie-breaking; using W={W_used}.")
        else:
            print(f"Objective: lexicographic tie-break with W={W_used}.")
        model.Maximize(W_used * utility10_var - penalty_var)
    else:
        model.Maximize(utility10_var - penalty_var)

    # --- Solve ---
    solver = cp_model.CpSolver()
    solver.parameters.num_search_workers = workers_used
    if time_limit_s is not None:
        solver.parameters.max_time_in_seconds = float(time_limit_s)

    if raw_cpsat_log:
        solver.parameters.log_search_progress = True
        solver.parameters.log_to_stdout = True
        try:
            solver.parameters.log_frequency_in_seconds = 1.0
        except Exception:
            pass

    t0 = time.perf_counter()
    cb = ProgressSummary(t0=t0, utility10_var=utility10_var, penalty_var=penalty_var) if progress_summary else None
    if cb is not None:
        # OR-Tools Python API compatibility:
        # - Some versions expose CpSolver.SolveWithSolutionCallback(model, cb)
        # - Newer versions use CpSolver.Solve(model, solution_callback=cb)
        try:
            status = solver.SolveWithSolutionCallback(model, cb)  # type: ignore[attr-defined]
        except AttributeError:
            try:
                status = solver.Solve(model, cb)  # positional callback
            except TypeError:
                status = solver.Solve(model, solution_callback=cb)
    else:
        status = solver.Solve(model)
    solve_elapsed_s = time.perf_counter() - t0

    # --- Final summary ---
    print("\nFinal summary:")
    print(f"  Solve time: {solve_elapsed_s:.3f}s (solver wall: {solver.WallTime():.3f}s)")
    print(f"  Status: {solver.StatusName(status)}")

    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        print("  No feasible solution found.")
        print(solver.ResponseStats())
        return

    util10 = int(solver.Value(utility10_var))
    pen = int(solver.Value(penalty_var))
    obj = int(round(solver.ObjectiveValue()))
    bound = int(round(solver.BestObjectiveBound()))
    gap = bound - obj

    print(f"  Best utility (LaTeX scale): {util10/10.0:.1f}")
    print(f"  Best penalty (overlap pairs): {pen}")
    if lexicographic_overlap_tiebreak:
        print(f"  Objective: maximize {W_used}*utility10 - penalty")
    else:
        print("  Objective: maximize utility10 - penalty")
    print(f"  Best bound (objective): {bound}")
    print(f"  Absolute gap (objective): {gap}")
    # NumSolutions() is not available on all OR-Tools Python wrappers.
    if hasattr(solver, "NumSolutions"):
        try:
            sol_count = int(solver.NumSolutions())  # type: ignore[attr-defined]
        except Exception:
            sol_count = None
    else:
        sol_count = None
    if sol_count is None and cb is not None:
        # Our callback tracks incumbents seen.
        sol_count = int(getattr(cb, "sols", 0))
    sol_s = "n/a" if sol_count is None else str(sol_count)
    print(f"  Conflicts: {solver.NumConflicts()}  Branches: {solver.NumBranches()}  Solutions: {sol_s}")

    # --- Write solution ---
    results = []
    m_scores, s_scores, ids = [], [], []
    for i in students:
        res_m = solver.Value(Main[i])
        res_s = solver.Value(Shadow[i])
        results.append({
            "Student": i,
            "Main": res_m,
            "M_Score": v[i][res_m],
            "Shadow": res_s,
            "S_Score": v[i][res_s],
            "Main_Part": "A" if solver.Value(p[res_m]) == 1 else "B",
            "Shadow_Part": "A" if solver.Value(p[res_s]) == 1 else "B",
            "Main_TopicCategory": cat[res_m],
            "Shadow_TopicCategory": cat[res_s],
        })
        m_scores.append(v[i][res_m])
        s_scores.append(v[i][res_s])
        ids.append(i)

    with open(filename, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=results[0].keys())
        writer.writeheader()
        writer.writerows(results)

    if plot:
        visualize_results(ids, m_scores, s_scores, n=n)




def solve_with_preferences_live(
    topic_titles: List[str],
    preferences: List[List[int]],
    M: int = 4,
    time_limit_s: float | None = 30,
    workers: int | None = None,
    lexicographic_overlap_tiebreak: bool = True,
    weight_W: int = 1000,
    progress_cb=None,
    stop_event=None,
) -> Dict[str, Any]:
    """Like solve_with_preferences, but streams progress and supports interruption."""
    n = len(topic_titles)
    if n == 0:
        raise ValueError("At least one topic is required")
    if any(len(row) != n for row in preferences):
        raise ValueError("preferences must be an n x n matrix")

    v = preferences
    r = list(range(n))
    topics = range(n)
    students = range(n)
    K, _ = choose_even_k(n, M)

    model = cp_model.CpModel()
    g = [model.NewBoolVar(f"g_{j}") for j in topics]
    p = [model.NewBoolVar(f"p_{j}") for j in topics]
    model.Add(p[1] == 1) if n >= 2 else model.Add(p[0] == 1)
    model.Add(sum(g) == K)

    u = []
    for j in topics:
        uj = model.NewBoolVar(f"u_{j}")
        model.Add(uj <= g[j])
        model.Add(uj <= p[j])
        model.Add(uj >= g[j] + p[j] - 1)
        u.append(uj)
    model.Add(sum(u) == K // 2)

    Main, Shadow, MainScore, ShadowScore = [], [], [], []
    main_is: Dict[Tuple[int, int], cp_model.IntVar] = {}
    shadow_is: Dict[Tuple[int, int], cp_model.IntVar] = {}

    for i in students:
        allowed = [j for j in topics if v[i][j] > 0]
        if len(allowed) < 2:
            raise ValueError(f"Student {i} has fewer than 2 non-veto topics (needs main != shadow).")

        main_i = model.NewIntVarFromDomain(cp_model.Domain.FromValues(allowed), f"Main_{i}")
        shadow_i = model.NewIntVarFromDomain(cp_model.Domain.FromValues(allowed), f"Shadow_{i}")
        Main.append(main_i)
        Shadow.append(shadow_i)
        model.Add(main_i != shadow_i)

        g_main_i = model.NewIntVar(0, 1, f"gMain_{i}")
        g_shadow_i = model.NewIntVar(0, 1, f"gShadow_{i}")
        model.AddElement(main_i, g, g_main_i)
        model.AddElement(shadow_i, g, g_shadow_i)
        model.Add(g_main_i == 1)
        model.Add(g_shadow_i == 1)

        p_main_i = model.NewIntVar(0, 1, f"pMain_{i}")
        p_shadow_i = model.NewIntVar(0, 1, f"pShadow_{i}")
        model.AddElement(main_i, p, p_main_i)
        model.AddElement(shadow_i, p, p_shadow_i)
        model.Add(p_main_i + p_shadow_i == 1)

        model.Add(main_i == r[i]).OnlyEnforceIf(g[r[i]])

        ms = model.NewIntVar(0, 5, f"MainScore_{i}")
        ss = model.NewIntVar(0, 5, f"ShadowScore_{i}")
        model.AddElement(main_i, v[i], ms)
        model.AddElement(shadow_i, v[i], ss)
        MainScore.append(ms)
        ShadowScore.append(ss)

        for j in allowed:
            b = model.NewBoolVar(f"main_is_{i}_{j}")
            model.Add(main_i == j).OnlyEnforceIf(b)
            model.Add(main_i != j).OnlyEnforceIf(b.Not())
            main_is[(i, j)] = b

            cvar = model.NewBoolVar(f"shadow_is_{i}_{j}")
            model.Add(shadow_i == j).OnlyEnforceIf(cvar)
            model.Add(shadow_i != j).OnlyEnforceIf(cvar.Not())
            shadow_is[(i, j)] = cvar

    for j in topics:
        main_count = sum(main_is[(i, j)] for i in students if (i, j) in main_is)
        shadow_count = sum(shadow_is[(i, j)] for i in students if (i, j) in shadow_is)
        model.Add(main_count <= M * g[j])
        model.Add(main_count >= (M - 1) * g[j])
        model.Add(shadow_count <= M * g[j])
        model.Add(shadow_count >= (M - 1) * g[j])

    overlap = []
    for i in students:
        for k in range(i + 1, n):
            eq_main = model.NewBoolVar(f"eqMain_{i}_{k}")
            eq_shadow = model.NewBoolVar(f"eqShadow_{i}_{k}")
            model.Add(Main[i] == Main[k]).OnlyEnforceIf(eq_main)
            model.Add(Main[i] != Main[k]).OnlyEnforceIf(eq_main.Not())
            model.Add(Shadow[i] == Shadow[k]).OnlyEnforceIf(eq_shadow)
            model.Add(Shadow[i] != Shadow[k]).OnlyEnforceIf(eq_shadow.Not())

            z = model.NewBoolVar(f"z_{i}_{k}")
            model.Add(z <= eq_main)
            model.Add(z <= eq_shadow)
            model.Add(z >= eq_main + eq_shadow - 1)
            overlap.append(z)

    utility10_var = model.NewIntVar(0, 150 * n, "utility10")
    model.Add(utility10_var == sum(20 * MainScore[i] + 10 * ShadowScore[i] for i in students))
    penalty_var = model.NewIntVar(0, n * (n - 1) // 2, "penalty")
    model.Add(penalty_var == sum(overlap))

    if lexicographic_overlap_tiebreak:
        min_safe_W = (n * (M - 1)) // 2 + 1
        W = max(weight_W, min_safe_W)
        model.Maximize(W * utility10_var - penalty_var)
    else:
        model.Maximize(utility10_var - penalty_var)

    solver = cp_model.CpSolver()
    solver.parameters.num_search_workers = multiprocessing.cpu_count() if workers is None else int(workers)
    if time_limit_s is not None:
        solver.parameters.max_time_in_seconds = float(time_limit_s)
    if progress_cb is not None:
        try:
            solver.parameters.log_search_progress = True
            solver.parameters.log_to_stdout = False
        except Exception:
            pass
        try:
            def _on_solver_log(line: str) -> None:
                msg = (line or "").strip()
                if not msg:
                    return
                low = msg.lower()
                if (
                    "solution" in low
                    or "objective" in low
                    or "bound" in low
                    or "gap" in low
                    or "status" in low
                    or "feasible" in low
                    or "optimal" in low
                ):
                    progress_cb(f"solver: {msg}")
            solver.log_callback = _on_solver_log  # type: ignore[attr-defined]
        except Exception:
            pass

    class LiveProgress(cp_model.CpSolverSolutionCallback):
        def __init__(self):
            super().__init__()
            self.sols = 0
            self.t0 = time.perf_counter()
            self.last_obj: float | None = None

        def OnSolutionCallback(self):
            self.sols += 1
            util = int(self.Value(utility10_var)) / 10.0
            pen = int(self.Value(penalty_var))
            elapsed = time.perf_counter() - self.t0

            try:
                obj = float(self.ObjectiveValue())
            except Exception:
                try:
                    obj = float(self.objective_value())
                except Exception:
                    obj = float("nan")

            try:
                bound = float(self.BestObjectiveBound())
            except Exception:
                try:
                    bound = float(self.best_objective_bound())
                except Exception:
                    bound = float("nan")

            gap = float("nan")
            if not math.isnan(obj) and not math.isnan(bound):
                gap = bound - obj

            def fmt_num(x: float) -> str:
                return "n/a" if math.isnan(x) else str(int(round(x)))

            if self.last_obj is None or (not math.isnan(obj) and obj > self.last_obj + 1e-9):
                self.last_obj = obj
            msg = (
                f"next solution found #{self.sols}: "
                f"t={elapsed:.2f}s util={util:.1f} pen={pen} "
                f"obj={fmt_num(obj)} bound={fmt_num(bound)} gap={fmt_num(gap)}"
            )
            if progress_cb is not None:
                progress_cb(msg)
            if stop_event is not None and stop_event.is_set():
                if progress_cb is not None:
                    progress_cb("interrupt requested; stopping search")
                self.StopSearch()

    cb = LiveProgress()
    if progress_cb is not None:
        try:
            proto = model.Proto()
            progress_cb(f"model stats: constraints={len(proto.constraints)} variables={len(proto.variables)}")
        except Exception:
            pass
        progress_cb("solve started")
    try:
        status = solver.SolveWithSolutionCallback(model, cb)  # type: ignore[attr-defined]
    except AttributeError:
        status = solver.Solve(model, cb)

    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        raise ValueError(f"No feasible solution found: {solver.StatusName(status)}")

    selected_topics = [j for j in topics if solver.Value(g[j]) == 1]
    assignments = []
    for i in students:
        main_topic = int(solver.Value(Main[i]))
        shadow_topic = int(solver.Value(Shadow[i]))
        assignments.append(
            {
                "student": i,
                "main_topic": main_topic,
                "main_title": topic_titles[main_topic],
                "main_score": int(v[i][main_topic]),
                "shadow_topic": shadow_topic,
                "shadow_title": topic_titles[shadow_topic],
                "shadow_score": int(v[i][shadow_topic]),
            }
        )

    overlaps = []
    for i in range(n):
        for k in range(i + 1, n):
            if assignments[i]["main_topic"] == assignments[k]["main_topic"] and assignments[i]["shadow_topic"] == assignments[k]["shadow_topic"]:
                overlaps.append((i, k))

    status_name = solver.StatusName(status)
    lowered = status_name.lower()
    if stop_event is not None and stop_event.is_set() and lowered == "feasible":
        status_name = "INTERRUPTED"
    if progress_cb is not None:
        progress_cb(f"solve ended with status={status_name}")

    return {
        "n": n,
        "K": K,
        "selected_topics": [
            {"id": j, "title": topic_titles[j], "partition": "A" if solver.Value(p[j]) == 1 else "B"}
            for j in selected_topics
        ],
        "assignments": assignments,
        "utility": int(solver.Value(utility10_var)) / 10.0,
        "penalty": int(solver.Value(penalty_var)),
        "status": status_name,
        "overlaps": overlaps,
    }
def solve_with_preferences(
    topic_titles: List[str],
    preferences: List[List[int]],
    M: int = 4,
    time_limit_s: float | None = 10,
    workers: int | None = None,
    lexicographic_overlap_tiebreak: bool = True,
    weight_W: int = 1000,
) -> Dict[str, Any]:
    """Solve assignment problem from explicit topics and voting matrix.

    Args:
        topic_titles: List of topic names, one per student/topic.
        preferences: n x n utility matrix in {0..5}.
        M: Group size upper bound.
    """
    n = len(topic_titles)
    if n == 0:
        raise ValueError("At least one topic is required")
    if any(len(row) != n for row in preferences):
        raise ValueError("preferences must be an n x n matrix")

    v = preferences
    r = list(range(n))
    topics = range(n)
    students = range(n)

    K, _ = choose_even_k(n, M)

    model = cp_model.CpModel()
    g = [model.NewBoolVar(f"g_{j}") for j in topics]
    p = [model.NewBoolVar(f"p_{j}") for j in topics]
    model.Add(p[1] == 1) if n >= 2 else model.Add(p[0] == 1)
    model.Add(sum(g) == K)

    u = []
    for j in topics:
        uj = model.NewBoolVar(f"u_{j}")
        model.Add(uj <= g[j])
        model.Add(uj <= p[j])
        model.Add(uj >= g[j] + p[j] - 1)
        u.append(uj)
    model.Add(sum(u) == K // 2)

    Main, Shadow, MainScore, ShadowScore = [], [], [], []
    main_is: Dict[Tuple[int, int], cp_model.IntVar] = {}
    shadow_is: Dict[Tuple[int, int], cp_model.IntVar] = {}

    for i in students:
        allowed = [j for j in topics if v[i][j] > 0]
        if len(allowed) < 2:
            raise ValueError(
                f"Student {i} has fewer than 2 non-veto topics (needs main != shadow)."
            )

        main_i = model.NewIntVarFromDomain(cp_model.Domain.FromValues(allowed), f"Main_{i}")
        shadow_i = model.NewIntVarFromDomain(cp_model.Domain.FromValues(allowed), f"Shadow_{i}")
        Main.append(main_i)
        Shadow.append(shadow_i)

        model.Add(main_i != shadow_i)

        g_main_i = model.NewIntVar(0, 1, f"gMain_{i}")
        g_shadow_i = model.NewIntVar(0, 1, f"gShadow_{i}")
        model.AddElement(main_i, g, g_main_i)
        model.AddElement(shadow_i, g, g_shadow_i)
        model.Add(g_main_i == 1)
        model.Add(g_shadow_i == 1)

        p_main_i = model.NewIntVar(0, 1, f"pMain_{i}")
        p_shadow_i = model.NewIntVar(0, 1, f"pShadow_{i}")
        model.AddElement(main_i, p, p_main_i)
        model.AddElement(shadow_i, p, p_shadow_i)
        model.Add(p_main_i + p_shadow_i == 1)

        model.Add(main_i == r[i]).OnlyEnforceIf(g[r[i]])

        ms = model.NewIntVar(0, 5, f"MainScore_{i}")
        ss = model.NewIntVar(0, 5, f"ShadowScore_{i}")
        model.AddElement(main_i, v[i], ms)
        model.AddElement(shadow_i, v[i], ss)
        MainScore.append(ms)
        ShadowScore.append(ss)

        for j in allowed:
            b = model.NewBoolVar(f"main_is_{i}_{j}")
            model.Add(main_i == j).OnlyEnforceIf(b)
            model.Add(main_i != j).OnlyEnforceIf(b.Not())
            main_is[(i, j)] = b

            cvar = model.NewBoolVar(f"shadow_is_{i}_{j}")
            model.Add(shadow_i == j).OnlyEnforceIf(cvar)
            model.Add(shadow_i != j).OnlyEnforceIf(cvar.Not())
            shadow_is[(i, j)] = cvar

    for j in topics:
        main_count = sum(main_is[(i, j)] for i in students if (i, j) in main_is)
        shadow_count = sum(shadow_is[(i, j)] for i in students if (i, j) in shadow_is)
        model.Add(main_count <= M * g[j])
        model.Add(main_count >= (M - 1) * g[j])
        model.Add(shadow_count <= M * g[j])
        model.Add(shadow_count >= (M - 1) * g[j])

    overlap = []
    for i in students:
        for k in range(i + 1, n):
            eq_main = model.NewBoolVar(f"eqMain_{i}_{k}")
            eq_shadow = model.NewBoolVar(f"eqShadow_{i}_{k}")
            model.Add(Main[i] == Main[k]).OnlyEnforceIf(eq_main)
            model.Add(Main[i] != Main[k]).OnlyEnforceIf(eq_main.Not())
            model.Add(Shadow[i] == Shadow[k]).OnlyEnforceIf(eq_shadow)
            model.Add(Shadow[i] != Shadow[k]).OnlyEnforceIf(eq_shadow.Not())

            z = model.NewBoolVar(f"z_{i}_{k}")
            model.Add(z <= eq_main)
            model.Add(z <= eq_shadow)
            model.Add(z >= eq_main + eq_shadow - 1)
            overlap.append(z)

    utility10_var = model.NewIntVar(0, 150 * n, "utility10")
    model.Add(utility10_var == sum(20 * MainScore[i] + 10 * ShadowScore[i] for i in students))
    penalty_var = model.NewIntVar(0, n * (n - 1) // 2, "penalty")
    model.Add(penalty_var == sum(overlap))

    if lexicographic_overlap_tiebreak:
        min_safe_W = (n * (M - 1)) // 2 + 1
        W = max(weight_W, min_safe_W)
        model.Maximize(W * utility10_var - penalty_var)
    else:
        model.Maximize(utility10_var - penalty_var)

    solver = cp_model.CpSolver()
    solver.parameters.num_search_workers = multiprocessing.cpu_count() if workers is None else int(workers)
    if time_limit_s is not None:
        solver.parameters.max_time_in_seconds = float(time_limit_s)

    status = solver.Solve(model)
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        raise ValueError(f"No feasible solution found: {solver.StatusName(status)}")

    selected_topics = [j for j in topics if solver.Value(g[j]) == 1]
    assignments = []
    for i in students:
        main_topic = int(solver.Value(Main[i]))
        shadow_topic = int(solver.Value(Shadow[i]))
        assignments.append(
            {
                "student": i,
                "main_topic": main_topic,
                "main_title": topic_titles[main_topic],
                "main_score": int(v[i][main_topic]),
                "shadow_topic": shadow_topic,
                "shadow_title": topic_titles[shadow_topic],
                "shadow_score": int(v[i][shadow_topic]),
            }
        )

    return {
        "n": n,
        "K": K,
        "selected_topics": [
            {"id": j, "title": topic_titles[j], "partition": "A" if solver.Value(p[j]) == 1 else "B"}
            for j in selected_topics
        ],
        "assignments": assignments,
        "utility": int(solver.Value(utility10_var)) / 10.0,
        "penalty": int(solver.Value(penalty_var)),
        "status": solver.StatusName(status),
    }


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Group assignment harness (OR-Tools CP-SAT).")
    parser.add_argument("--n", type=int, default=40, help="Number of students/topics.")
    parser.add_argument("--M", type=int, default=4, help="Max group size (min is M-1).")
    parser.add_argument("--C", type=int, default=8, help="Number of topic categories (used in category modes).")
    parser.add_argument("--seed", type=int, default=42, help="RNG seed.")
    parser.add_argument(
        "--pref_mode",
        choices=["category_uniform", "category", "category_mode3", "random"],
        default="category_uniform",
        help=(
            "Preference generator: "
            "'category_uniform' (real category utilities + veto + quintile binning), "
            "'category' (favorite-based), "
            "'category_mode3' (peaked at 3), "
            "or 'random' (i.i.d. raw scores)."
        ),
    )
    parser.add_argument(
        "--time_limit_s",
        type=float,
        default=None,
        help="Solver time limit (seconds). Omit for no limit.",
    )
    parser.add_argument(
        "--workers",
        type=int,
        default=None,
        help="Number of CP-SAT search workers (threads). Default: use all available logical cores.",
    )
    parser.add_argument("--W", type=int, default=1000, help="Weight for strict tie-break objective (if enabled).")
    parser.add_argument(
        "--no_lex",
        action="store_true",
        help="Disable strict lexicographic tie-break (use utility10 - penalty instead).",
    )
    parser.add_argument("--no_plot", action="store_true", help="Disable plotting.")
    parser.add_argument("--out", type=str, default="assignments_ortools.csv", help="Output CSV filename.")

    parser.add_argument(
        "--no_progress",
        action="store_true",
        help="Disable the compact progress summary (new incumbents).",
    )
    parser.add_argument(
        "--raw_cpsat_log",
        action="store_true",
        help="Enable the built-in CP-SAT search log (verbose).",
    )

    args = parser.parse_args()

    solve_lab_ortools(
        n=args.n,
        M=args.M,
        C=args.C,
        filename=args.out,
        time_limit_s=args.time_limit_s,
        seed=args.seed,
        pref_mode=args.pref_mode,
        lexicographic_overlap_tiebreak=not args.no_lex,
        weight_W=args.W,
        plot=not args.no_plot,
        progress_summary=not args.no_progress,
        raw_cpsat_log=args.raw_cpsat_log,
        workers=args.workers,
    )
