# -*- coding: utf-8 -*-
"""
HB 스크립트
- 단계: preference -> recommend (given pref=1) -> intent (given rec=1) -> purchase (given intent=1)
- 랜덤효과: 세그먼트, 충성도 (+ 옵션: 세그×충성도 상호작용)
- 산출: 단계별 파라미터 요약, 그룹별 추정 전환율, 체인전환율, 간단 Sankey 집계
"""

import numpy as np
import pandas as pd
import pymc as pm
import arviz as az

# =========================
# 0) 유틸
# =========================
def sigmoid(x):
    return 1 / (1 + np.exp(-x))

def q025(x): return float(np.quantile(x, 0.025))
def q975(x): return float(np.quantile(x, 0.975))

def encode_factors(df, seg_col="segment", loy_col="loyalty"):
    df = df.copy()
    df["seg_code"] = df[seg_col].astype("category").cat.codes
    df["loy_code"] = df[loy_col].astype("category").cat.codes
    n_seg = df["seg_code"].nunique()
    n_loy = df["loy_code"].nunique()
    seg_levels = list(df[seg_col].astype("category").cat.categories)
    loy_levels = list(df[loy_col].astype("category").cat.categories)
    return df, n_seg, n_loy, seg_levels, loy_levels

def fit_stage_hier_logit(df, ycol, mask, use_interaction=False, seed=42):
    """
    df: seg_code, loy_code, ycol(0/1) 포함
    mask: True/False 마스크 (조건부 학습 대상)
    use_interaction: 세그×충성도 상호작용 랜덤효과 추가
    """
    idx = np.where(mask)[0]
    if len(idx) == 0:
        raise ValueError(f"[{ycol}] 학습 대상이 없습니다(마스크 0건).")

    y_obs = df.loc[idx, ycol].astype(int).values
    seg_idx = df.loc[idx, "seg_code"].values
    loy_idx = df.loc[idx, "loy_code"].values

    n_seg = df["seg_code"].nunique()
    n_loy = df["loy_code"].nunique()

    with pm.Model() as model:
        mu = pm.Normal("mu", 0.0, 1.0)

        sigma_seg = pm.HalfNormal("sigma_seg", 1.0)
        sigma_loy = pm.HalfNormal("sigma_loy", 1.0)
        a_seg = pm.Normal("a_seg", 0, sigma_seg, shape=n_seg)
        a_loy = pm.Normal("a_loy", 0, sigma_loy, shape=n_loy)

        if use_interaction:
            sigma_int = pm.HalfNormal("sigma_int", 1.0)
            # 상호작용: seg x loy
            a_int = pm.Normal("a_int", 0, sigma_int, shape=(n_seg, n_loy))
            logit_p = mu + a_seg[seg_idx] + a_loy[loy_idx] + a_int[seg_idx, loy_idx]
        else:
            logit_p = mu + a_seg[seg_idx] + a_loy[loy_idx]

        p = pm.Deterministic("p", pm.math.sigmoid(logit_p))
        pm.Bernoulli("y_like", p=p, observed=y_obs)

        trace = pm.sample(1000, tune=1000, target_accept=0.9, chains=2, random_seed=seed)

    return trace, idx

def posterior_group_rates(df, trace, idx, seg_col="segment", loy_col="loyalty"):
    post_p = trace.posterior["p"].stack(draws=("chain", "draw")).mean("draws").values  # len(idx)
    out = df[[seg_col, loy_col]].copy()
    out["p_hat"] = np.nan
    out.loc[idx, "p_hat"] = post_p
    group = out.groupby([seg_col, loy_col], dropna=False)["p_hat"].mean().reset_index()
    return group

def chain_rate(pref, rec, intent, buy):
    return pref * rec * intent * buy

def sankey_counts(n_base, pref, rec, intent, buy):
    pref_users   = int(round(pref   * n_base))
    rec_users    = int(round(rec    * pref_users))
    intent_users = int(round(intent * rec_users))
    buy_users    = int(round(buy    * intent_users))
    return dict(
        pref_users=pref_users,
        rec_users=rec_users,
        intent_users=intent_users,
        buy_users=buy_users
    )

# =========================
# 1) 예시 데이터 생성
# =========================
np.random.seed(7)
N = 500
segments = np.random.choice(["A","B","C"], size=N, p=[0.4,0.4,0.2])
loyalty  = np.random.choice(["Low","Mid","High"], size=N, p=[0.3,0.5,0.2])

# 진짜 확률(숨겨진)
base = -0.2
seg_eff = {"A":0.0, "B":0.4, "C":-0.2}
loy_eff = {"Low":-0.3, "Mid":0.0, "High":0.3}

# 단계별 로짓(간단 예시) – 실제는 단계별 다른 영향도도 가능
def make_prob(seg, loy, add_bias=0.0):
    z = base + seg_eff[seg] + loy_eff[loy] + add_bias
    return sigmoid(z)

p_pref   = np.array([make_prob(s, l, add_bias=0.0)  for s, l in zip(segments, loyalty)])
y_pref   = np.random.binomial(1, p_pref)

# recommend 는 pref=1인 사람만 대상
p_rec    = np.array([make_prob(s, l, add_bias=0.1)  for s, l in zip(segments, loyalty)])
y_rec    = np.where(y_pref==1, np.random.binomial(1, p_rec), 0)

# intent 는 rec=1인 사람만 대상
p_intent = np.array([make_prob(s, l, add_bias=0.2)  for s, l in zip(segments, loyalty)])
y_intent = np.where(y_rec==1, np.random.binomial(1, p_intent), 0)

# purchase 는 intent=1인 사람만 대상
p_buy    = np.array([make_prob(s, l, add_bias=0.3)  for s, l in zip(segments, loyalty)])
y_buy    = np.where(y_intent==1, np.random.binomial(1, p_buy), 0)

df = pd.DataFrame({
    "segment": segments,
    "loyalty": loyalty,
    "pref": y_pref,
    "rec": y_rec,
    "intent": y_intent,
    "buy": y_buy
})

# =========================
# 2) 인코딩
# =========================
df, n_seg, n_loy, seg_levels, loy_levels = encode_factors(df)

# =========================
# 3) 단계별 학습 (조건부 마스크)
# =========================
use_interaction = True  # 세그×충성도 상호작용 랜덤효과 사용여부

# stage 1: preference (전체)
mask_pref = np.ones(len(df), dtype=bool)
trace_pref, idx_pref = fit_stage_hier_logit(df, "pref", mask_pref, use_interaction=use_interaction, seed=11)

# stage 2: recommend | pref=1
mask_rec = (df["pref"] == 1).values
trace_rec, idx_rec = fit_stage_hier_logit(df, "rec", mask_rec, use_interaction=use_interaction, seed=22)

# stage 3: intent | rec=1
mask_intent = (df["rec"] == 1).values
trace_intent, idx_intent = fit_stage_hier_logit(df, "intent", mask_intent, use_interaction=use_interaction, seed=33)

# stage 4: purchase | intent=1
mask_buy = (df["intent"] == 1).values
trace_buy, idx_buy = fit_stage_hier_logit(df, "buy", mask_buy, use_interaction=use_interaction, seed=44)

# =========================
# 4) 요약 및 그룹별 전환율
# =========================
print("\n=== [Summary] Hyper-parameters ===")
print("Preference:\n", az.summary(trace_pref, var_names=["mu","sigma_seg","sigma_loy","sigma_int"] if use_interaction else ["mu","sigma_seg","sigma_loy"]))
print("Recommend:\n",  az.summary(trace_rec,  var_names=["mu","sigma_seg","sigma_loy","sigma_int"] if use_interaction else ["mu","sigma_seg","sigma_loy"]))
print("Intent:\n",     az.summary(trace_intent,  var_names=["mu","sigma_seg","sigma_loy","sigma_int"] if use_interaction else ["mu","sigma_seg","sigma_loy"]))
print("Purchase:\n",   az.summary(trace_buy,  var_names=["mu","sigma_seg","sigma_loy","sigma_int"] if use_interaction else ["mu","sigma_seg","sigma_loy"]))

g_pref   = posterior_group_rates(df, trace_pref,   idx_pref)
g_rec    = posterior_group_rates(df, trace_rec,    idx_rec)
g_intent = posterior_group_rates(df, trace_intent, idx_intent)
g_buy    = posterior_group_rates(df, trace_buy,    idx_buy)

g_pref.rename(columns={"p_hat":"pref_hat"}, inplace=True)
g_rec.rename(columns={"p_hat":"rec_hat"}, inplace=True)
g_intent.rename(columns={"p_hat":"intent_hat"}, inplace=True)
g_buy.rename(columns={"p_hat":"buy_hat"}, inplace=True)

# 병합하여 그룹별 체인전환율까지
grp = g_pref.merge(g_rec, on=["segment","loyalty"], how="left") \
            .merge(g_intent, on=["segment","loyalty"], how="left") \
            .merge(g_buy, on=["segment","loyalty"], how="left")

grp["chain_hat"] = grp[["pref_hat","rec_hat","intent_hat","buy_hat"]].prod(axis=1)

print("\n=== [Group-level posterior mean rates] ===")
print(grp.sort_values(["segment","loyalty"]).to_string(index=False))

# 전체 평균(무게 없는 단순 평균)
overall_pref   = float(grp["pref_hat"].mean())
overall_rec    = float(grp["rec_hat"].mean())
overall_intent = float(grp["intent_hat"].mean())
overall_buy    = float(grp["buy_hat"].mean())
overall_chain  = float(grp["chain_hat"].mean())

print("\n=== [Overall mean rates (unweighted)] ===")
print(f"pref={overall_pref:.3f}, rec={overall_rec:.3f}, intent={overall_intent:.3f}, buy={overall_buy:.3f}, chain={overall_chain:.3f}")

# =========================
# 5) 전체 N 기준 기대 카운트
# =========================
N_total = len(df)
sk = sankey_counts(
    n_base=N_total,
    pref=overall_pref,
    rec=overall_rec,
    intent=overall_intent,
    buy=overall_buy
)
print("\n=== [Sankey expected counts @ overall] ===")
print(sk)

# =========================
# 6) (선택) 엑셀로 저장
# =========================
try:
    out_xlsx = "bayes_funnel_results.xlsx"
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as w:
        # 단계별 하이퍼파라미터 요약
        az.summary(trace_pref).to_csv("trace_pref_summary.csv")   # 참고: CSV도 함께 남김(선택)
        az.summary(trace_rec).to_csv("trace_rec_summary.csv")
        az.summary(trace_intent).to_csv("trace_intent_summary.csv")
        az.summary(trace_buy).to_csv("trace_buy_summary.csv")

        # 그룹별 결과
        grp.to_excel(w, sheet_name="group_rates", index=False)

        # 전체요약
        pd.DataFrame([{
            "pref": overall_pref,
            "rec": overall_rec,
            "intent": overall_intent,
            "buy": overall_buy,
            "chain": overall_chain,
            "N_total": N_total,
            **{f"sankey_{k}": v for k, v in sk.items()}
        }]).to_excel(w, sheet_name="overall_summary", index=False)

    print(f"\n✅ 저장 완료: {out_xlsx}")
except Exception as e:
    print(f"\n(참고) 엑셀 저장 생략: {e}")
