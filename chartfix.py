def _generate_matplotlib_charts(kpis: dict, table_stats: dict) -> dict:
    """Generates Base64 PNG charts locally using Python matplotlib (100% offline)."""
    charts = {"overall_b64": "", "table_b64": ""}

    # ---------------- CHART 1: DOUGHNUT CHART (OVERALL BREAKDOWN) ----------------
    try:
        fig1, ax1 = plt.subplots(figsize=(4.5, 2.6), dpi=120)
        matched = kpis.get("matched_rows", 0)
        mismatched = kpis.get("mismatched_rows", 0)
        exceptions = kpis.get("approved_exceptions", 0)

        vals = [matched, mismatched, exceptions]
        labels = [f"Matched ({matched})", f"Mismatch ({mismatched})", f"25.4 Exc ({exceptions})"]
        colors = ["#28a745", "#dc3545", "#17a2b8"]

        if sum(vals) > 0:
            wedges, _ = ax1.pie(
                vals, colors=colors, startangle=140,
                wedgeprops=dict(width=0.42, edgecolor="white", linewidth=2)
            )
            ax1.legend(wedges, labels, loc="center left", bbox_to_anchor=(0.92, 0.5), frameon=False, fontsize=8.5)
        else:
            ax1.text(0.5, 0.5, "No Data", ha="center", va="center", fontsize=11, color="#718096")
            ax1.axis("off")

        plt.tight_layout()
        buf1 = io.BytesIO()
        fig1.savefig(buf1, format="png", transparent=True, bbox_inches="tight")
        plt.close(fig1)
        buf1.seek(0)
        charts["overall_b64"] = "data:image/png;base64," + base64.b64encode(buf1.read()).decode("utf-8")
    except Exception as e:
        print(f"[IBC Matplotlib Doughnut Error] {e}")

    # ---------------- CHART 2: GROUPED BAR CHART (BY CC&B TABLE) ----------------
    try:
        fig2, ax2 = plt.subplots(figsize=(6.8, 2.6), dpi=120)
        tbl_names = list(table_stats.keys())
        short_names = [t.replace("CI_", "") for t in tbl_names]
        m_vals = [table_stats[t]["match"] for t in tbl_names]
        mm_vals = [table_stats[t]["mismatch"] for t in tbl_names]

        x = range(len(short_names))
        width = 0.36

        bars1 = ax2.bar([i - width / 2 for i in x], m_vals, width, label="Match", color="#28a745")
        bars2 = ax2.bar([i + width / 2 for i in x], mm_vals, width, label="Mismatch", color="#dc3545")

        ax2.set_xticks(list(x))
        ax2.set_xticklabels(short_names, fontsize=8, fontweight="bold", color="#2d3748")
        ax2.tick_params(axis="y", labelsize=8)
        ax2.spines["top"].set_visible(False)
        ax2.spines["right"].set_visible(False)
        ax2.grid(axis="y", linestyle="--", alpha=0.4)
        ax2.legend(frameon=False, fontsize=8.5, loc="upper right")

        # Add value labels on top of bars
        for bar in list(bars1) + list(bars2):
            h = bar.get_height()
            if h > 0:
                ax2.annotate(
                    f"{int(h)}",
                    xy=(bar.get_x() + bar.get_width() / 2, h),
                    xytext=(0, 2), textcoords="offset points",
                    ha="center", va="bottom", fontsize=7.5, fontweight="bold"
                )

        plt.tight_layout()
        buf2 = io.BytesIO()
        fig2.savefig(buf2, format="png", transparent=True, bbox_inches="tight")
        plt.close(fig2)
        buf2.seek(0)
        charts["table_b64"] = "data:image/png;base64," + base64.b64encode(buf2.read()).decode("utf-8")
    except Exception as e:
        print(f"[IBC Matplotlib Bar Error] {e}")

    return charts