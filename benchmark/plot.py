#!/usr/bin/env python3
"""Generate benchmark comparison chart for rbxl README."""

import matplotlib.pyplot as plt
import matplotlib
import numpy as np

matplotlib.use("Agg")

# Benchmark data (5000 rows x 10 cols)
categories = ["write", "read", "read\nvalues_only"]

data = {
    "rbxl + native":    [0.038, 0.084, 0.027],
    "rbxl (pure Ruby)": [0.098, 0.315, 0.247],
    "openpyxl":         [0.398, 0.344, 0.288],
    "caxlsx":           [0.484, None,  None],
    "roo":              [None,  1.078, None],
    "rubyXL":           [2.185, 2.028, None],
}

# None = library does not support this mode
# (roo is read-only, caxlsx is write-only)

colors = {
    "rbxl + native":    "#2563eb",
    "rbxl (pure Ruby)": "#60a5fa",
    "openpyxl":         "#f59e0b",
    "caxlsx":           "#a855f7",
    "roo":              "#10b981",
    "rubyXL":           "#ef4444",
}

# Footnotes for missing data
notes = {
    "caxlsx": "write only",
    "roo":    "read only, no values_only",
    "rubyXL": "no values_only",
}

libs = list(data.keys())
max_y = 0.55
bar_width = 0.12

fig, ax = plt.subplots(figsize=(10, 5.5))

# For each category, only place bars that have data, centered
for cat_idx, cat in enumerate(categories):
    active = [(lib, data[lib][cat_idx]) for lib in libs if data[lib][cat_idx] is not None]
    n = len(active)
    offsets = np.arange(n) - (n - 1) / 2

    for i, (lib, val) in enumerate(active):
        x = cat_idx + offsets[i] * bar_width
        display_val = min(val, max_y)
        ax.bar(x, display_val, bar_width * 0.9, color=colors[lib], label=lib)

        if val > max_y:
            ax.text(x, max_y * 0.88, f"{val:.1f}s",
                    ha="center", va="top", fontsize=8, fontweight="bold", color="white")
        else:
            ax.text(x, val + 0.008, f"{val:.3f}",
                    ha="center", va="bottom", fontsize=7.5)

ax.set_ylabel("Time (seconds)", fontsize=11)
ax.set_title("rbxl benchmark — 5,000 rows × 10 columns", fontsize=13, fontweight="bold")
ax.set_xticks(range(len(categories)))
ax.set_xticklabels(categories, fontsize=11)
ax.set_ylim(0, max_y)
ax.grid(axis="y", alpha=0.3)

# Deduplicate legend
handles, labels = ax.get_legend_handles_labels()
seen = {}
unique_handles, unique_labels = [], []
for h, l in zip(handles, labels):
    if l not in seen:
        seen[l] = True
        # Append note if library has missing modes
        display = f"{l}  ({notes[l]})" if l in notes else l
        unique_handles.append(h)
        unique_labels.append(display)

ax.legend(unique_handles, unique_labels, loc="upper right", fontsize=8.5)

fig.tight_layout()
fig.savefig("benchmark/chart.png", dpi=150)
print("Saved benchmark/chart.png")
