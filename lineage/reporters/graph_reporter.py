"""Graph reporter for Excel Lineage Detector - hierarchical PNG diagram."""

from __future__ import annotations
from pathlib import Path

from lineage.models import DataConnection

# Color palette by category.
# When adding a new extractor category, add entries to all three dicts below.
# Unknown categories fall back to DEFAULT_COLOR / DEFAULT_BG automatically.
CATEGORY_COLORS = {
    "database":   "#1565C0",
    "powerquery": "#6A1B9A",
    "file":       "#2E7D32",
    "web":        "#E65100",
    "hyperlink":  "#AD1457",
    "vba":        "#880E4F",
    "pivot":      "#37474F",
    "formula":    "#00695C",
    "metadata":   "#4E342E",
    "ole":        "#4527A0",
    "input":      "#F57F17",   # amber/gold - signals manual data entry
}
CATEGORY_BG = {
    "database":   "#E3F2FD",
    "powerquery": "#F3E5F5",
    "file":       "#E8F5E9",
    "web":        "#FFF3E0",
    "hyperlink":  "#FCE4EC",
    "vba":        "#FCE4EC",
    "pivot":      "#ECEFF1",
    "formula":    "#E0F7FA",
    "metadata":   "#EFEBE9",
    "ole":        "#EDE7F6",
    "input":      "#FFF9C4",   # light yellow
}
# Controls display order in the graph (left column, top to bottom).
# New categories not listed here are appended alphabetically at the end.
CATEGORY_ORDER = [
    "database", "powerquery", "file", "web",
    "hyperlink", "vba", "formula", "pivot", "metadata", "ole", "input",
]
DEFAULT_COLOR = "#546E7A"
DEFAULT_BG    = "#ECEFF1"
TARGET_COLOR  = "#E65100"
TARGET_BG     = "#FFF8E1"
TARGET_BORDER = "#BF360C"

# Layout constants (all in inches; 1 data unit == 1 inch)
LEFT_MARGIN  = 0.55
RIGHT_MARGIN = 0.40
TOP_MARGIN   = 0.65
BOT_MARGIN   = 0.45
NODE_W       = 4.60   # source node width
TARGET_W     = 2.80   # target node width
H_GAP        = 1.50   # horizontal space between source right edge and target left edge
FONT_SIZE    = 8.5
LINE_H       = 0.175  # height per text line
NODE_PAD_X   = 0.14
NODE_PAD_Y   = 0.11
NODE_GAP     = 0.12   # vertical gap between nodes in same category
CAT_HDR_H    = 0.30   # category header strip height
CAT_GAP      = 0.28   # extra vertical gap between categories
WRAP_CHARS   = 46     # chars per line inside a source node


class GraphReporter:
    """Generates a hierarchical PNG lineage diagram."""

    def write(
        self,
        connections: list[DataConnection],
        input_path: Path,
        out_dir: Path,
    ) -> Path:
        try:
            import matplotlib
            matplotlib.use("Agg")
            import matplotlib.pyplot as plt
            import matplotlib.patches as mpatches
            from matplotlib.patches import FancyBboxPatch
            import textwrap
            from collections import defaultdict
        except ImportError as e:
            raise ImportError(f"matplotlib is required: {e}")

        stem = input_path.stem
        out  = out_dir / f"{stem}_lineage_graph.png"

        # ── 1. Deduplicate and group by category ─────────────────────────
        seen: dict[tuple, DataConnection] = {}
        for c in connections:
            key = (c.category, c.source)
            if key not in seen:
                seen[key] = c

        cat_sources: dict[str, list[DataConnection]] = defaultdict(list)
        for (cat, _src), conn in seen.items():
            cat_sources[cat].append(conn)

        ordered_cats = [c for c in CATEGORY_ORDER if c in cat_sources]
        ordered_cats += sorted(c for c in cat_sources if c not in ordered_cats)

        if not ordered_cats:
            return self._empty(input_path, out, stem)

        # ── 2. Calculate y positions (y=0 at top, increases downward) ────
        nodes: list[dict] = []
        cat_bands: list[dict] = []
        y = TOP_MARGIN

        for cat in ordered_cats:
            srcs   = cat_sources[cat]
            color  = CATEGORY_COLORS.get(cat, DEFAULT_COLOR)
            bg     = CATEGORY_BG.get(cat, DEFAULT_BG)

            band_y_start = y
            y += CAT_HDR_H  # space for the category label

            for conn in srcs:
                raw = conn.source or conn.sub_type or conn.id
                main_lines = textwrap.wrap(raw, width=WRAP_CHARS) or [raw[:WRAP_CHARS]]
                n_lines    = len(main_lines) + 1          # +1 for the [sub_type] footer
                h          = NODE_PAD_Y * 2 + LINE_H * n_lines

                nodes.append({
                    "x":        LEFT_MARGIN,
                    "y":        y,
                    "w":        NODE_W,
                    "h":        h,
                    "cx":       LEFT_MARGIN + NODE_W / 2,
                    "cy":       y + h / 2,
                    "main":     main_lines,
                    "sub_type": conn.sub_type,
                    "category": cat,
                    "color":    color,
                })
                y += h + NODE_GAP

            # shrink the last NODE_GAP (avoid extra space before cat gap)
            y -= NODE_GAP
            cat_bands.append({
                "cat":     cat,
                "y_start": band_y_start,
                "y_end":   y,
                "color":   color,
                "bg":      bg,
            })
            y += CAT_GAP

        y += BOT_MARGIN
        total_h = y

        # ── 3. Target node ────────────────────────────────────────────────
        tgt_x      = LEFT_MARGIN + NODE_W + H_GAP
        tgt_lines  = textwrap.wrap(stem, width=20) or [stem[:20]]
        tgt_lines += ["(target)"]
        tgt_h      = NODE_PAD_Y * 2 + LINE_H * len(tgt_lines)
        tgt_y      = total_h / 2 - tgt_h / 2
        target     = {
            "x": tgt_x, "y": tgt_y, "w": TARGET_W, "h": tgt_h,
            "cx": tgt_x + TARGET_W / 2,
            "cy": tgt_y + tgt_h / 2,
            "lines": tgt_lines,
        }

        fig_w = LEFT_MARGIN + NODE_W + H_GAP + TARGET_W + RIGHT_MARGIN
        fig_h = max(total_h, 5.0)

        # ── 4. Draw ───────────────────────────────────────────────────────
        fig = plt.figure(figsize=(fig_w, fig_h), facecolor="white")
        # full-figure axes; data units == inches
        ax = fig.add_axes([0, 0, 1, 1])
        ax.set_xlim(0, fig_w)
        ax.set_ylim(fig_h, 0)   # y=0 at top
        ax.set_facecolor("white")
        ax.axis("off")

        # ── Category background bands ─────────────────────────────────────
        for band in cat_bands:
            bx = LEFT_MARGIN - 0.10
            bw = NODE_W + 0.20
            bh = band["y_end"] - band["y_start"]

            ax.add_patch(FancyBboxPatch(
                (bx, band["y_start"]), bw, bh,
                boxstyle="round,pad=0.06",
                facecolor=band["bg"],
                edgecolor=band["color"],
                linewidth=0.7,
                alpha=0.55,
                zorder=1,
            ))
            # Category label centred in the header strip
            ax.text(
                LEFT_MARGIN - 0.02,
                band["y_start"] + CAT_HDR_H / 2,
                band["cat"].upper(),
                fontsize=7.5, fontweight="bold",
                color=band["color"],
                va="center", ha="left",
                zorder=3,
            )

        # ── Horizontal "bus" connector line ───────────────────────────────
        bus_x   = LEFT_MARGIN + NODE_W + H_GAP * 0.42
        bus_y_t = nodes[0]["cy"]
        bus_y_b = nodes[-1]["cy"]
        ax.plot(
            [bus_x, bus_x], [bus_y_t, bus_y_b],
            color="#CFD8DC", linewidth=1.5, zorder=2,
        )

        # ── Arrows: source → bus → target ────────────────────────────────
        for node in nodes:
            src_right = node["x"] + node["w"]
            src_cy    = node["cy"]
            clr       = node["color"] + "90"  # 56% opacity hex

            # horizontal stub to bus
            ax.plot(
                [src_right, bus_x], [src_cy, src_cy],
                color=clr, linewidth=1.0, zorder=2,
            )
            # dot on the bus
            ax.plot(bus_x, src_cy, "o",
                    color=node["color"], markersize=3, zorder=3)

        # single arrow from bus midpoint to target left edge
        bus_mid_y = (bus_y_t + bus_y_b) / 2
        ax.annotate(
            "",
            xy=(target["x"], target["cy"]),
            xytext=(bus_x, bus_mid_y),
            arrowprops=dict(
                arrowstyle="-|>",
                color="#78909C",
                lw=1.8,
                mutation_scale=14,
            ),
            zorder=4,
        )

        # ── Source nodes ──────────────────────────────────────────────────
        for node in nodes:
            ax.add_patch(FancyBboxPatch(
                (node["x"], node["y"]), node["w"], node["h"],
                boxstyle="round,pad=0.06",
                facecolor=node["color"],
                edgecolor="white",
                linewidth=1.4,
                zorder=4,
            ))

            # Main label lines (bold white)
            main_text = "\n".join(node["main"])
            n_main    = len(node["main"])
            # Centre of main-text block sits above the sub_type line
            main_cy = node["y"] + NODE_PAD_Y + (n_main * LINE_H) / 2
            ax.text(
                node["cx"], main_cy,
                main_text,
                fontsize=FONT_SIZE, fontweight="bold",
                color="white", va="center", ha="center", ma="center",
                linespacing=1.25,
                zorder=5,
            )

            # Sub-type footer (lighter, smaller)
            sub_y = node["y"] + node["h"] - NODE_PAD_Y - LINE_H * 0.50
            ax.text(
                node["cx"], sub_y,
                f"[{node['sub_type']}]",
                fontsize=FONT_SIZE - 1.5,
                color="white", alpha=0.82,
                va="center", ha="center",
                zorder=5,
            )

        # ── Target node ───────────────────────────────────────────────────
        ax.add_patch(FancyBboxPatch(
            (target["x"], target["y"]), target["w"], target["h"],
            boxstyle="round,pad=0.10",
            facecolor=TARGET_BG,
            edgecolor=TARGET_BORDER,
            linewidth=2.2,
            zorder=4,
        ))
        ax.text(
            target["cx"], target["cy"],
            "\n".join(target["lines"]),
            fontsize=10, fontweight="bold",
            color="#1A1A1A", va="center", ha="center", ma="center",
            linespacing=1.3,
            zorder=5,
        )

        # ── Title ─────────────────────────────────────────────────────────
        ax.text(
            fig_w / 2, TOP_MARGIN * 0.38,
            f"Data Lineage  ·  {input_path.name}",
            fontsize=11, fontweight="bold", color="#212121",
            va="center", ha="center", zorder=6,
        )
        ax.text(
            fig_w / 2, TOP_MARGIN * 0.72,
            f"{len(connections)} connections  ·  {len(seen)} unique sources",
            fontsize=8.5, color="#616161",
            va="center", ha="center", zorder=6,
        )

        # ── Legend ────────────────────────────────────────────────────────
        legend_cats = [c for c in CATEGORY_ORDER if c in cat_sources]
        legend_cats += [c for c in cat_sources if c not in legend_cats]
        patches = [
            mpatches.Patch(facecolor=TARGET_BG, edgecolor=TARGET_BORDER,
                           linewidth=1.2, label="Target file"),
        ]
        for cat in legend_cats:
            patches.append(mpatches.Patch(
                color=CATEGORY_COLORS.get(cat, DEFAULT_COLOR),
                label=cat.title(),
            ))

        leg = ax.legend(
            handles=patches,
            loc="lower right",
            bbox_to_anchor=(fig_w - 0.08, fig_h - 0.08),
            bbox_transform=ax.transData,
            fontsize=7.5,
            framealpha=0.95,
            edgecolor="#B0BEC5",
            ncol=1,
        )
        leg.get_frame().set_linewidth(0.8)

        # ── Save ──────────────────────────────────────────────────────────
        fig.savefig(str(out), dpi=150, bbox_inches="tight", facecolor="white")
        plt.close(fig)
        return out

    # ── helpers ──────────────────────────────────────────────────────────

    def _empty(self, input_path: Path, out: Path, stem: str) -> Path:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        fig, ax = plt.subplots(figsize=(8, 4))
        ax.text(0.5, 0.5, f"No connections found in\n{input_path.name}",
                ha="center", va="center", fontsize=13, transform=ax.transAxes)
        ax.axis("off")
        fig.savefig(str(out), dpi=150, bbox_inches="tight", facecolor="white")
        plt.close(fig)
        return out
