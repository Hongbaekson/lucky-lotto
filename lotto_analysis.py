"""6/45 비복원 추출을 기준으로 한 진단과 시간순 확률예측 비교."""

from collections import Counter
from itertools import combinations
from math import ceil, comb, sqrt
import random

import numpy as np
from scipy.stats import binomtest, chisquare, fisher_exact, hypergeom, norm

import lotto


VERSION = "weekly-1"
P = 6 / 45
WARMUP = 208
HOLDOUT = 208
PRIOR = 52
BOOTSTRAPS = 4000
BLOCK = 13
MODELS = {"uniform": "균등 확률", "all": "누적 빈도", "recent26": "최근 26회",
          "recent52": "최근 52회", "recent104": "최근 104회", "ema26": "최근 가중 평균",
          "transition": "직전 번호와 다음 회차 관계"}


def indicators(draws):
    matrix = np.zeros((len(draws), 45), dtype=np.float64)
    for i, draw in enumerate(draws):
        matrix[i, np.asarray(draw.numbers) - 1] = 1
    return matrix


def holm(pvalues):
    """검정 사이의 독립성을 요구하지 않는 Holm 다중검정 보정."""
    pvalues = np.asarray(pvalues, dtype=float)
    order = np.argsort(pvalues, kind="stable")
    adjusted = np.empty(len(order))
    adjusted[order] = np.minimum(1, np.maximum.accumulate(pvalues[order] * np.arange(len(order), 0, -1)))
    return adjusted.tolist()


def diagnostics(draws):
    if len(draws) < WARMUP + 2 * HOLDOUT:
        raise lotto.LottoError("심층 분석에는 1회부터 최소 624회의 이력이 필요합니다.")
    matrix = indicators(draws)
    n = len(draws)
    counts = matrix.sum(axis=0).astype(int)
    rows = []
    distributions = []

    def add(group, name, observed, expected, pvalue, method):
        rows.append({"group": group, "name": name, "observed": float(observed),
                     "expected": float(expected), "p": float(pvalue), "method": method})

    for number, count in enumerate(counts, 1):
        add("번호 빈도", str(number), count, n * P, binomtest(int(count), n, P).pvalue, "양측 정확 이항검정")
    pairs = Counter(pair for draw in draws for pair in combinations(draw.numbers, 2))
    for a, b in combinations(range(1, 46), 2):
        count = pairs[a, b]
        add("번호쌍 빈도", f"{a:02}-{b:02}", count, n / 66,
            binomtest(count, n, 1 / 66).pvalue, "양측 정확 이항검정; 한 쌍의 출현 확률 1/66")
    split = n // 2
    first = matrix[:split].sum(axis=0).astype(int)
    second = counts - first
    for i in range(45):
        pvalue = fisher_exact([[int(first[i]), split - int(first[i])],
                              [int(second[i]), n - split - int(second[i])]]).pvalue
        add("전후반 변화", str(i + 1), second[i] / (n - split) - first[i] / split,
            0, pvalue, "전반/후반 출현 비율 차이; 양측 Fisher 정확검정")
    overlap_variance = float(hypergeom.var(45, 6, 6))
    for lag in range(1, 11):
        overlaps = (matrix[:-lag] * matrix[lag:]).sum(axis=1)
        z = (overlaps.mean() - 0.8) / sqrt(overlap_variance / len(overlaps))
        add("회차 간 반복", f"{lag}회 간격", overlaps.mean(), 0.8, 2 * norm.sf(abs(z)),
            "비복원 추출의 평균 중복 개수; 큰 표본 정규 근사")

    def distribution(label, observed, expected, labels):
        observed, expected = np.asarray(observed), np.asarray(expected, dtype=float)
        if np.any(expected < 5):
            raise lotto.LottoError(f"{label}: 근사 검정에 필요한 기대 표본 수가 부족합니다.")
        result = chisquare(observed, expected * observed.sum() / expected.sum())
        add("조합 분포", label, result.statistic, len(observed) - 1, result.pvalue,
            "Pearson 카이제곱; 관측값=검정통계량, 기준값=자유도")
        for key, actual, theoretical in zip(labels, observed, expected):
            distributions.append({"group": label, "bin": str(key), "observed": int(actual),
                                  "expected": float(theoretical)})

    odd = matrix[:, ::2].sum(axis=1).astype(int)
    distribution("홀수 개수", np.bincount(odd, minlength=7),
                 hypergeom.pmf(np.arange(7), 45, 23, 6) * n, range(7))
    # 45개 중 6개를 고르는 모든 조합의 합 분포를 동적 계획으로 정확히 계산한다.
    ways = np.zeros((7, 271), dtype=np.int64)
    ways[0, 0] = 1
    for number in range(1, 46):
        for count in range(6, 0, -1):
            ways[count, number:] += ways[count - 1, :-number]
    bins = [21, 90, 110, 130, 150, 170, 190, 256]
    sums = matrix @ np.arange(1, 46)
    observed, _ = np.histogram(sums, bins=bins)
    expected = [ways[6, a:b].sum() / comb(45, 6) * n for a, b in zip(bins, bins[1:])]
    distribution("번호 합계", observed, expected, [f"{a}~{b - 1}" for a, b in zip(bins, bins[1:])])
    adjacent = [sum(b == a + 1 for a, b in zip(d.numbers, d.numbers[1:])) for d in draws]
    adjacency_prob = [comb(5, k) * comb(40, 6 - k) / comb(45, 6) for k in range(6)]
    distribution("연속 번호쌍 개수", np.bincount(np.minimum(adjacent, 3), minlength=4),
                 np.array([*adjacency_prob[:3], sum(adjacency_prob[3:])]) * n, ["0", "1", "2", "3 이상"])
    for row, adjusted in zip(rows, holm([row["p"] for row in rows])):
        row["adjusted_p"] = adjusted
        row["significant"] = adjusted < 0.05
    gaps = [b.drawn_at - a.drawn_at for a, b in zip(draws, draws[1:])]
    return {"draws": n, "tests": rows, "distributions": distributions,
            "significant_count": sum(row["significant"] for row in rows),
            "number_expected": n * P, "pair_expected": n / 66,
            "sum_mean": float(sums.mean()), "sum_expected": 138,
            "overlap_mean": float((matrix[:-1] * matrix[1:]).sum(axis=1).mean()),
            "duplicate_combinations": n - len({d.numbers for d in draws}),
            "nonweekly_dates": sum(gap.days != 7 for gap in gaps)}


def forecasts(matrix):
    """i행을 보기 전에 앞의 i개 회차만으로 예측하고, 마지막에 다음 회차를 예측한다."""
    cumulative = np.vstack([np.zeros(45), matrix.cumsum(axis=0)])
    transition = np.zeros((45, 45))
    ema, mass = np.zeros(45), 0.0
    decay = 0.5 ** (1 / 26)
    for i in range(len(matrix) + 1):
        if i >= WARMUP:
            probabilities = {"uniform": np.full(45, P), "all": (cumulative[i] + PRIOR * P) / (i + PRIOR)}
            for window in (26, 52, 104):
                used = min(i, window)
                probabilities[f"recent{window}"] = (cumulative[i] - cumulative[i - used] + PRIOR * P) / (used + PRIOR)
            probabilities["ema26"] = (ema + PRIOR * P) / (mass + PRIOR)
            previous = np.flatnonzero(matrix[i - 1])
            probabilities["transition"] = ((transition[previous] + PRIOR * P) /
                                            (cumulative[i - 1, previous, None] + PRIOR)).mean(axis=0)
            yield i, probabilities
        if i < len(matrix):
            if i:
                transition[np.ix_(np.flatnonzero(matrix[i - 1]), np.flatnonzero(matrix[i]))] += 1
            ema = decay * ema + matrix[i]
            mass = decay * mass + 1


def bootstrap_gain(values, seed):
    """양수=균등 예측보다 Brier 손실 감소. 13회 블록 부트스트랩 근사 구간."""
    values = np.asarray(values, dtype=float)
    rng = np.random.default_rng(seed)
    starts = rng.integers(0, len(values), size=(BOOTSTRAPS, ceil(len(values) / BLOCK)))
    indices = (starts[:, :, None] + np.arange(BLOCK)) % len(values)
    sampled_means = values[indices.reshape(BOOTSTRAPS, -1)[:, :len(values)]].mean(axis=1)
    mean = float(values.mean())
    low, high = np.quantile(sampled_means, [0.025, 0.975])
    pvalue = (1 + np.count_nonzero(sampled_means - mean >= mean)) / (BOOTSTRAPS + 1)
    return {"gain": mean, "ci_low": float(low), "ci_high": float(high), "p": float(pvalue)}


def sample_two(probabilities, rng):
    weights = np.asarray(probabilities, dtype=float)
    if weights.shape != (45,) or not np.isfinite(weights).all() or np.any(weights <= 0):
        raise lotto.LottoError("추천 가중치가 유효하지 않습니다.")
    tickets = []
    for _ in range(200):
        if np.all(weights == weights[0]):
            ticket = sorted(rng.sample(range(1, 46), 6))
            if ticket not in tickets:
                tickets.append(ticket)
            if len(tickets) == 2:
                return tickets
            continue
        available, remaining, chosen = list(range(1, 46)), weights.tolist(), []
        for _ in range(6):
            number = rng.choices(available, weights=remaining, k=1)[0]
            index = available.index(number)
            available.pop(index)
            remaining.pop(index)
            chosen.append(number)
        ticket = sorted(chosen)
        if ticket not in tickets:
            tickets.append(ticket)
        if len(tickets) == 2:
            return tickets
    raise lotto.LottoError("서로 다른 2게임 생성에 실패했습니다.")


def evaluate(draws, log=print):
    if len(draws) < WARMUP + 2 * HOLDOUT:
        raise lotto.LottoError("주간 분석에는 최소 624회의 이력이 필요합니다.")
    matrix = indicators(draws)
    cutoff = len(draws) - HOLDOUT
    errors = {model: [] for model in MODELS}
    log_losses = {model: [] for model in MODELS}
    illustrative = {model: Counter() for model in MODELS}
    for index, predictions in forecasts(matrix):
        if index == len(draws):
            next_probabilities = {model: probabilities.tolist() for model, probabilities in predictions.items()}
            break
        actual = matrix[index]
        for model, probabilities in predictions.items():
            errors[model].append(float(np.mean((probabilities - actual) ** 2)))
            log_losses[model].append(float(-np.mean(actual * np.log(probabilities) + (1 - actual) * np.log1p(-probabilities))))
            if index >= cutoff:
                rng = random.Random(f"{VERSION}:illustration:{draws[index].round}:{model}")
                for ticket in sample_two(probabilities, rng):
                    count, prize = lotto.rank(ticket, draws[index])
                    illustrative[model][f"rank{prize}"] += 1
                    illustrative[model]["hits"] += count
    split = cutoff - WARMUP
    baseline = np.asarray(errors["uniform"])
    results = []
    for index, model in enumerate(MODELS):
        loss = np.asarray(errors[model])
        gain = baseline - loss
        validation = bootstrap_gain(gain[:split], 20260929 + index)
        test = bootstrap_gain(gain[split:], 20261029 + index)
        results.append({"model": model, "label": MODELS[model],
                        "validation_brier": float(loss[:split].mean()),
                        "test_brier": float(loss[split:].mean()),
                        "test_log_loss": float(np.mean(log_losses[model][split:])),
                        "validation": validation, "test": test,
                        "illustrative_mean_hits": illustrative[model]["hits"] / (HOLDOUT * 2),
                        "illustrative_ranks": {str(rank): illustrative[model][f"rank{rank}"] for rank in range(1, 6)}})
    for result, adjusted in zip(results[1:], holm([r["validation"]["p"] for r in results[1:]])):
        result["validation"]["adjusted_p"] = adjusted
    results[0]["validation"]["adjusted_p"] = 1.0
    # 후보는 선택 구간에서 한 번만 결정하며, 마지막 208회 성적으로 다른 후보를 고르지 않는다.
    candidate = min(results[1:], key=lambda row: row["validation_brier"])
    passed = (candidate["validation"]["ci_low"] > 0 and candidate["validation"]["adjusted_p"] < 0.05
              and candidate["test"]["ci_low"] > 0 and candidate["test"]["p"] < 0.05)
    selected = candidate["model"] if passed else "uniform"
    reason = ("사전 선택 후보가 분리 평가에서도 평균 확률오차 개선 기준을 통과했습니다. 당첨 조합 예측력이 입증된 것은 아닙니다."
              if passed else "분리 평가까지 일관된 개선을 확인하지 못해 균등 확률로 2게임을 생성합니다.")
    log(f"모델 비교 완료: {MODELS[selected]}. {reason}")
    return {"models": results, "candidate": candidate["model"], "selected": selected, "reason": reason,
            "validation_start": draws[WARMUP].round, "validation_end": draws[cutoff - 1].round,
            "test_start": draws[cutoff].round, "test_end": draws[-1].round,
            "bootstrap_samples": BOOTSTRAPS, "bootstrap_block": BLOCK,
            "next_probabilities": next_probabilities}


def analyze(draws, log=print):
    log("번호 45개·번호쌍 990개·전후반 변화·1~10회 간격·조합 분포를 검정합니다.")
    findings = diagnostics(draws)
    log(f"진단 {len(findings['tests'])}개 완료. 전체 Holm 보정 후 유의한 항목: {findings['significant_count']}개")
    log("확률예측 7개를 시간순으로 비교합니다. 마지막 208회는 모델 선택에서 제외합니다.")
    return {"diagnostics": findings, "evaluation": evaluate(draws, log)}
