"""Regenerate four-model KCI plots from copied CSVs; no training or Office edits."""
from pathlib import Path
import hashlib
import json
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

RUN = Path(__file__).resolve().parents[1]/"outputs/chat-replay-20261028"
DATA = RUN/"inputs/data"
OUT = RUN/"plots/four-model"
OUT.mkdir(exist_ok=True, parents=True)
MODELS = ["No-Gain", "Gain", "Gain + SS=0.01", "Difficult-GradNorm ratio=0.01"]
LABELS = ["No-Gain", "Gain (0.05)", "Gain + SS (0.01)", r"DG $\rho=0.01$"]
STYLES = [("#333333", "s", "--"), ("#0072B2", "o", "-"), ("#E69F00", "^", "-."), ("#6A3D9A", "P", "--")]
PAIRS = ["air2 → H₂S", "T2 → H₂S", "air2 → SO₂", "T2 → SO₂"]
PAIR_LABELS = [r"Air2 $\rightarrow$ H$_2$S", r"T2 $\rightarrow$ H$_2$S", r"Air2 $\rightarrow$ SO$_2$", r"T2 $\rightarrow$ SO$_2$"]
DISTS = ["in_range", "soft_ood"]
plt.rcParams.update({"font.family": "Arial", "font.size": 10, "pdf.fonttype": 42})
audits = []


def finish(fig, axes, name, source, count):
    handles, labels = axes.flat[0].get_legend_handles_labels()
    assert labels == LABELS
    fig.legend(handles, labels, loc="upper center", ncol=4, frameon=False, fontsize=10)
    assert all(not a.get_title() for a in axes.flat)
    assert not fig._suptitle
    for ext in ("png", "pdf"):
        fig.savefig(OUT/f"{name}.{ext}", dpi=600, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    audits.append({"figure": name, "models": MODELS, "source": str(source),
                   "sha256": hashlib.sha256(source.read_bytes()).hexdigest(),
                   "csv_values_checked": count, "titles": False, "error_bars": False})


source = DATA/"seven_model_all18_kci_values.csv"
d = pd.read_csv(source)
d = d[d.model.isin(MODELS)].copy()
assert len(d) == 32
np.testing.assert_allclose(100*(d.tp+d.tn)/d.valid_comparisons, d.kci_percent, rtol=1e-12)
fig, axes = plt.subplots(1, 2, figsize=(11, 4.5), squeeze=False, sharey=True)
for k, dist in enumerate(DISTS):
    ax = axes[0, k]
    for model, label, (color, marker, ls) in zip(MODELS, LABELS, STYLES):
        values = []
        for target, mv in [("B35_H2S", "air2_SP"), ("B35_H2S", "HEATER2_output_T_SP"), ("B35_SO2", "air2_SP"), ("B35_SO2", "HEATER2_output_T_SP")]:
            r = d[(d.model == model)&(d.distribution == dist)&(d.target == target)&(d.mv == mv)]
            assert len(r) == 1
            values.append(float(r.iloc[0].kci_percent))
        line, = ax.plot(range(4), values, color=color, marker=marker, ls=ls, label=label)
        np.testing.assert_array_equal(line.get_ydata(), values)
    ax.set_xticks(range(4), PAIR_LABELS, rotation=15, ha="right")
    ax.set_ylim(0, 103)
    ax.set_ylabel("KCI (%)")
    ax.text(.01, 1.02, f"({chr(97+k)})", transform=ax.transAxes, va="bottom", weight="bold")
    ax.grid(axis="y", alpha=.25)
fig.subplots_adjust(top=.83, bottom=.23, left=.08, right=.98, wspace=.2)
finish(fig, axes, "four_model_all18_kci", source, len(d))

for column, filename, name in [
    ("teacher_sign_label", "conditional_kci_by_teacher_sign.csv", "four_model_conditional_kci_by_teacher_sign"),
    ("magnitude_bin", "conditional_kci_by_gain_magnitude.csv", "four_model_conditional_kci_by_gain_magnitude")]:
    source = DATA/filename
    d = pd.read_csv(source)
    d = d[d.model.isin(MODELS)].copy()
    groups = list(dict.fromkeys(d[column]))
    assert len(groups) == (2 if "sign" in column else 4)
    assert len(d) == 4*2*4*len(groups)
    valid = d.comparisons > 0
    np.testing.assert_allclose(100*d.loc[valid, "correct"]/d.loc[valid, "comparisons"], d.loc[valid, "kci_percent"], rtol=1e-12)
    assert d.loc[~valid, "kci_percent"].isna().all()
    fig, axes = plt.subplots(2, 4, figsize=(12.6, 6), sharey=True)
    for i, dist in enumerate(DISTS):
        for j, pair in enumerate(PAIRS):
            ax = axes[i, j]
            for model, label, (color, marker, ls) in zip(MODELS, LABELS, STYLES):
                r = d[(d.model == model)&(d.distribution == dist)&(d.pair == pair)]
                assert len(r) == len(groups), (pair, dist, model)
                values = r.set_index(column).loc[groups, "kci_percent"].to_numpy(float)
                line, = ax.plot(range(len(groups)), values, color=color, marker=marker, ls=ls, label=label)
                np.testing.assert_allclose(line.get_ydata(), values, rtol=0, atol=0, equal_nan=True)
            ax.set_xticks(range(len(groups)), groups)
            ax.set_xlim(-.2, len(groups)-.8)
            ax.tick_params(axis="x", labelsize=8)
            ax.set_ylim(0, 103)
            ax.text(.01, 1.02, f"({chr(97+i*4+j)})", transform=ax.transAxes, va="bottom", weight="bold")
            if j == 0:
                ax.set_ylabel("KCI (%)")
            if j == 2 and "sign" in column:
                empty = d[(d.distribution == dist)&(d.pair == pair)&(d.comparisons == 0)][column].unique()
                for group in empty:
                    ax.text(groups.index(group), 8, "N/A", ha="center", color="#666666")
            ax.grid(axis="y", alpha=.25)
    fig.subplots_adjust(top=.84, bottom=.08, left=.07, right=.99, hspace=.4, wspace=.25)
    finish(fig, axes, name, source, len(d))
(OUT/"kci-replot-audit.json").write_text(json.dumps(audits, indent=2), encoding="utf-8")
print("KCI replot passed:", sum(a["csv_values_checked"] for a in audits), "CSV values across", len(audits), "figures")
