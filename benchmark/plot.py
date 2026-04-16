#!/usr/bin/env python3
"""Generate benchmark comparison chart for rbxl README."""

from datetime import datetime
from pathlib import Path
import re

import matplotlib.pyplot as plt
import matplotlib
import numpy as np

matplotlib.use("Agg")

# Benchmark data (5000 rows x 10 cols)
categories = ["write", "read", "read\nvalues_only"]

data = {
    "rbxl + native":        [0.046, 0.087, 0.039],
    "rbxl (pure Ruby)":     [0.084, 0.291, 0.219],
    "fast_excel":           [0.195, None, None],
    "fast_excel constant":  [0.116, None, None],
    "excelize":             [0.149, 0.143, None],
    "rust_xlsxwriter":      [0.627, None, None],
    "calamine":             [None, 0.490, None],
    "exceljs":              [0.082, 0.202, None],
    "sheetjs":              [0.157, 0.211, None],
    "openpyxl":             [0.356, 0.204, 0.185],
}

# None = library does not support this mode
# (roo is read-only, caxlsx is write-only)

colors = {
    "rbxl + native":        "#2563eb",
    "rbxl (pure Ruby)":     "#60a5fa",
    "fast_excel":           "#7c3aed",
    "fast_excel constant":  "#a78bfa",
    "excelize":             "#0f766e",
    "rust_xlsxwriter":      "#b45309",
    "calamine":             "#92400e",
    "exceljs":              "#f97316",
    "sheetjs":              "#22c55e",
    "openpyxl":             "#f59e0b",
}

# Footnotes for missing data
notes = {
    "fast_excel": "write-only benchmark",
    "fast_excel constant": "write-only, constant_memory: true",
    "excelize": "no values_only benchmark",
    "rust_xlsxwriter": "write-only benchmark",
    "calamine": "read-only benchmark",
    "exceljs": "no values_only benchmark",
    "sheetjs": "no values_only benchmark",
}

libs = list(data.keys())
max_y = 0.68
bar_width = 0.085

fig, ax = plt.subplots(figsize=(13.2, 6.4))

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
output_dir = Path("benchmark")
readme_path = Path("README.md")
timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
latest_path = output_dir / "chart.png"
timestamped_path = output_dir / f"chart-{timestamp}.png"

for old_chart in output_dir.glob("chart-*.png"):
    old_chart.unlink()

fig.savefig(latest_path, dpi=150)
fig.savefig(timestamped_path, dpi=150)

readme_text = readme_path.read_text()
updated_readme = re.sub(r"!\[Benchmark chart\]\(benchmark/chart(?:-\d{8}-\d{6})?\.png\)",
                        f"![Benchmark chart]({timestamped_path.as_posix()})",
                        readme_text)
if updated_readme != readme_text:
    readme_path.write_text(updated_readme)

print(f"Saved {latest_path}")
print(f"Saved {timestamped_path}")
print(f"Updated {readme_path}")
