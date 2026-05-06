from __future__ import annotations

from pathlib import Path
import textwrap

import nbformat as nbf


ROOT = Path(__file__).resolve().parent
NOTEBOOK = ROOT / "rhodea_garlic_peel_analysis.ipynb"


def md(text: str):
    return nbf.v4.new_markdown_cell(textwrap.dedent(text).strip())


def code(text: str):
    return nbf.v4.new_code_cell(textwrap.dedent(text).strip())


nb = nbf.v4.new_notebook()
nb["metadata"] = {
    "kernelspec": {
        "display_name": "Anaconda Python",
        "language": "python",
        "name": "python3",
    },
    "language_info": {"name": "python", "pygments_lexer": "ipython3"},
}

cells = [
    md(
        """
        # Antimicrobial Potential of Ilocos White Garlic Peel Powder

        This notebook analyzes `Book-of-Analysis-2026-rhodea.xlsx` using the objectives and methods described in `reference.docx`.

        **Objectives translated into computation**

        1. Organize the full factorial drying design for garlic peel powder: temperature (`50`, `60`, `70` °C) and drying time (`4`, `6`, `8` hours).
        2. Summarize physicochemical responses: water activity (`Aw`), moisture content (`%MC`), color (`L*`, `a*`, `b*`, derived `C*`, `h°`, and `Delta_E`), water absorption capacity (`WAC`), and water solubility index (`WSI`).
        3. Evaluate antimicrobial activity against `Staphylococcus aureus`, `Escherichia coli`, and `Salmonella spp.` using zone of inhibition values.
        4. Apply the stated methods: two-way ANOVA for physicochemical responses, one-way ANOVA for antimicrobial activity by treatment, and Response Surface Methodology (RSM) to visualize and identify optimized processing conditions.

        **Source note**

        The document states the design uses drying times of `4`, `6`, and `8` hours, but its treatment matrix lists Treatment 9 as `70°C, 7 hrs`. Because this is otherwise a 3x3 full factorial design, this notebook treats Treatment 9 as `70°C, 8 hrs` and keeps this note visible for manuscript correction.
        """
    ),
    code(
        """
        from pathlib import Path
        import json
        import math
        import warnings

        import numpy as np
        import pandas as pd
        import matplotlib.pyplot as plt
        import seaborn as sns
        import statsmodels.api as sm
        import statsmodels.formula.api as smf
        from statsmodels.stats.anova import anova_lm
        from scipy import stats

        warnings.filterwarnings("ignore", category=FutureWarning)
        pd.set_option("display.max_columns", 80)
        pd.set_option("display.width", 160)

        ROOT = Path.cwd()
        XLSX_PATH = ROOT / "Book-of-Analysis-2026-rhodea.xlsx"
        FIG_DIR = ROOT / "figures"
        OUT_DIR = ROOT / "outputs"
        FIG_DIR.mkdir(exist_ok=True)
        OUT_DIR.mkdir(exist_ok=True)

        sns.set_theme(style="whitegrid", context="notebook")
        plt.rcParams["figure.dpi"] = 130
        plt.rcParams["savefig.dpi"] = 200
        """
    ),
    md(
        """
        ## 1. Experimental Design Map

        The workbook stores observations by treatment number. The map below supplies the independent variables needed for ANOVA and RSM.
        """
    ),
    code(
        """
        treatment_map = pd.DataFrame(
            {
                "Treatment": range(1, 10),
                "Temperature_C": [50, 50, 50, 60, 60, 60, 70, 70, 70],
                "Time_h": [4, 6, 8, 4, 6, 8, 4, 6, 8],
            }
        )
        treatment_map
        """
    ),
    md(
        """
        ## 2. Load and Clean the Workbook Data

        The workbook contains a combined physicochemical sheet plus separate antimicrobial mini-tables. This section converts those layouts into analysis-ready long tables.
        """
    ),
    code(
        """
        def load_physicochemical_data(path: Path) -> pd.DataFrame:
            raw = pd.read_excel(path, sheet_name="DO NOT INPUT ANYTHING HERE!!!", header=None)
            df = raw.iloc[2:, :13].copy()
            df.columns = [
                "Treatment", "WAC", "L_star", "a_star", "b_star", "C_workbook", "H_workbook",
                "Delta_E_workbook", "Aw", "MC_percent", "WSI", "pH", "TSS"
            ]
            df = df[pd.to_numeric(df["Treatment"], errors="coerce").notna()].copy()
            df["Treatment"] = df["Treatment"].astype(int)
            for col in df.columns.drop("Treatment"):
                df[col] = pd.to_numeric(df[col], errors="coerce")

            df = df.merge(treatment_map, on="Treatment", how="left")
            df["Replicate"] = df.groupby("Treatment").cumcount() + 1

            # Recompute derived color metrics because workbook C, H, and Delta_E columns are all zero.
            df["Chroma_C"] = np.sqrt(df["a_star"] ** 2 + df["b_star"] ** 2)
            df["Hue_deg"] = np.degrees(np.arctan2(df["b_star"], df["a_star"]))
            baseline = df[df["Treatment"] == 1][["L_star", "a_star", "b_star"]].mean()
            df["Delta_E_vs_T1"] = np.sqrt(
                (df["L_star"] - baseline["L_star"]) ** 2
                + (df["a_star"] - baseline["a_star"]) ** 2
                + (df["b_star"] - baseline["b_star"]) ** 2
            )
            return df


        def load_antimicrobial_data(path: Path) -> pd.DataFrame:
            raw = pd.read_excel(path, sheet_name="Zone inihibition", header=None)
            blocks = {
                "Staphylococcus aureus": 0,
                "Escherichia coli": 8,
                "Salmonella spp.": 16,
            }
            rows = []
            for organism, start_col in blocks.items():
                sub = raw.iloc[2:, start_col:start_col + 2].copy()
                sub.columns = ["Treatment", "Zone_mm"]
                sub["Treatment"] = pd.to_numeric(sub["Treatment"], errors="coerce").ffill()
                sub["Zone_mm"] = pd.to_numeric(sub["Zone_mm"], errors="coerce")
                sub = sub[sub["Treatment"].notna()].copy()
                # The workbook formulas treat a blank E. coli replicate as zero.
                # This is retained here for consistency with the workbook means.
                sub["Zone_mm"] = sub["Zone_mm"].fillna(0)
                sub["Treatment"] = sub["Treatment"].astype(int)
                sub["Organism"] = organism
                sub["Replicate"] = sub.groupby(["Organism", "Treatment"]).cumcount() + 1
                rows.append(sub)
            df = pd.concat(rows, ignore_index=True)
            df = df[df["Treatment"].between(0, 9)].copy()
            df = df.merge(treatment_map, on="Treatment", how="left")
            return df


        phys = load_physicochemical_data(XLSX_PATH)
        anti = load_antimicrobial_data(XLSX_PATH)

        print("Physicochemical rows:", phys.shape)
        print("Antimicrobial rows:", anti.shape)
        display(phys.head())
        display(anti.head())
        """
    ),
    code(
        """
        phys.to_csv(OUT_DIR / "clean_physicochemical_data.csv", index=False)
        anti.to_csv(OUT_DIR / "clean_antimicrobial_data.csv", index=False)
        """
    ),
    md("## 3. Descriptive Statistics"),
    code(
        """
        phys_responses = ["WAC", "L_star", "a_star", "b_star", "Chroma_C", "Hue_deg", "Delta_E_vs_T1", "Aw", "MC_percent", "WSI"]
        phys_summary = (
            phys.groupby(["Treatment", "Temperature_C", "Time_h"])[phys_responses]
            .agg(["mean", "std", "count"])
            .round(4)
        )
        phys_summary.to_csv(OUT_DIR / "physicochemical_summary_by_treatment.csv")
        phys_summary
        """
    ),
    code(
        """
        anti_summary = (
            anti[anti["Treatment"].between(1, 9)]
            .groupby(["Organism", "Treatment", "Temperature_C", "Time_h"])["Zone_mm"]
            .agg(["mean", "std", "count"])
            .round(4)
            .reset_index()
        )
        anti_summary.to_csv(OUT_DIR / "antimicrobial_summary_by_treatment.csv", index=False)
        anti_summary
        """
    ),
    md("## 4. Physicochemical Visualizations"),
    code(
        """
        plot_responses = ["WAC", "Aw", "MC_percent", "WSI", "L_star", "a_star", "b_star", "Chroma_C"]
        means = phys.groupby(["Temperature_C", "Time_h"])[plot_responses].mean().reset_index()

        fig, axes = plt.subplots(2, 4, figsize=(15, 7), constrained_layout=True)
        for ax, response in zip(axes.ravel(), plot_responses):
            pivot = means.pivot(index="Temperature_C", columns="Time_h", values=response)
            sns.heatmap(pivot, annot=True, fmt=".2f", cmap="viridis", ax=ax, cbar=False)
            ax.set_title(response)
            ax.set_xlabel("Time (h)")
            ax.set_ylabel("Temperature (°C)")
        fig.suptitle("Treatment Mean Heatmaps for Physicochemical Responses", y=1.03, fontsize=14)
        heatmap_path = FIG_DIR / "physicochemical_heatmaps.png"
        fig.savefig(heatmap_path, bbox_inches="tight")
        plt.show()
        heatmap_path
        """
    ),
    code(
        """
        fig, axes = plt.subplots(2, 2, figsize=(11, 8), constrained_layout=True)
        for ax, response in zip(axes.ravel(), ["WAC", "Aw", "MC_percent", "WSI"]):
            sns.pointplot(
                data=phys,
                x="Time_h", y=response, hue="Temperature_C",
                dodge=True, errorbar="sd", markers="o", capsize=.08, ax=ax
            )
            ax.set_title(f"Interaction Plot: {response}")
            ax.set_xlabel("Time (h)")
            ax.legend(title="Temp (°C)")
        interaction_path = FIG_DIR / "physicochemical_interaction_plots.png"
        fig.savefig(interaction_path, bbox_inches="tight")
        plt.show()
        interaction_path
        """
    ),
    md(
        """
        ## 5. Two-Way ANOVA for Physicochemical Responses

        Model used for each response:

        `response ~ C(Temperature_C) + C(Time_h) + C(Temperature_C):C(Time_h)`

        This matches the methodology's two-way ANOVA for independent variables temperature and time.
        """
    ),
    code(
        """
        def two_way_anova(df: pd.DataFrame, response: str) -> pd.DataFrame:
            data = df[["Temperature_C", "Time_h", response]].dropna().copy()
            model = smf.ols(f"Q('{response}') ~ C(Temperature_C) * C(Time_h)", data=data).fit()
            table = anova_lm(model, typ=2).reset_index().rename(columns={"index": "Source"})
            table.insert(0, "Response", response)
            return table

        anova_phys = pd.concat([two_way_anova(phys, r) for r in phys_responses], ignore_index=True)
        anova_phys["Significant_0.05"] = anova_phys["PR(>F)"] < 0.05
        anova_phys_rounded = anova_phys.round({"sum_sq": 5, "df": 0, "F": 4, "PR(>F)": 5})
        anova_phys_rounded.to_csv(OUT_DIR / "two_way_anova_physicochemical.csv", index=False)
        anova_phys_rounded
        """
    ),
    code(
        """
        significant_phys = (
            anova_phys[anova_phys["Source"] != "Residual"]
            .assign(p_value=lambda d: d["PR(>F)"])
            .query("p_value < 0.05")
            [["Response", "Source", "F", "p_value"]]
            .sort_values(["Response", "p_value"])
        )
        significant_phys.round(5)
        """
    ),
    md("## 6. One-Way ANOVA for Antimicrobial Activity"),
    code(
        """
        def one_way_anova_by_organism(df: pd.DataFrame) -> pd.DataFrame:
            rows = []
            for organism, sub in df[df["Treatment"].between(1, 9)].groupby("Organism"):
                groups = [g["Zone_mm"].values for _, g in sub.groupby("Treatment")]
                f_stat, p_value = stats.f_oneway(*groups)
                rows.append({"Organism": organism, "F": f_stat, "p_value": p_value, "Significant_0.05": p_value < 0.05})
            return pd.DataFrame(rows)

        anti_anova = one_way_anova_by_organism(anti).round(5)
        anti_anova.to_csv(OUT_DIR / "one_way_anova_antimicrobial.csv", index=False)
        anti_anova
        """
    ),
    code(
        """
        fig, ax = plt.subplots(figsize=(10, 5), constrained_layout=True)
        sns.barplot(
            data=anti[anti["Treatment"].between(1, 9)],
            x="Treatment", y="Zone_mm", hue="Organism",
            errorbar="sd", capsize=.08, ax=ax
        )
        ax.set_title("Zone of Inhibition by Treatment and Organism")
        ax.set_ylabel("Zone of inhibition (mm)")
        anti_plot_path = FIG_DIR / "antimicrobial_zone_by_treatment.png"
        fig.savefig(anti_plot_path, bbox_inches="tight")
        plt.show()
        anti_plot_path
        """
    ),
    md(
        """
        ## 7. Response Surface Methodology (RSM)

        A second-order response surface is fitted on treatment means:

        `response ~ Temperature + Time + Temperature² + Time² + Temperature×Time`

        This is used for visualization and optimization screening.
        """
    ),
    code(
        """
        def fit_rsm(mean_df: pd.DataFrame, response: str):
            d = mean_df[["Temperature_C", "Time_h", response]].dropna().copy()
            d["Temp2"] = d["Temperature_C"] ** 2
            d["Time2"] = d["Time_h"] ** 2
            d["Temp_Time"] = d["Temperature_C"] * d["Time_h"]
            model = smf.ols(f"Q('{response}') ~ Temperature_C + Time_h + Temp2 + Time2 + Temp_Time", data=d).fit()
            return model

        phys_means = phys.groupby(["Treatment", "Temperature_C", "Time_h"], as_index=False)[phys_responses].mean()
        anti_mean_wide = (
            anti[anti["Treatment"].between(1, 9)]
            .groupby(["Treatment", "Temperature_C", "Time_h", "Organism"], as_index=False)["Zone_mm"].mean()
            .pivot(index=["Treatment", "Temperature_C", "Time_h"], columns="Organism", values="Zone_mm")
            .reset_index()
            .rename(columns={
                "Staphylococcus aureus": "Zone_Staph",
                "Escherichia coli": "Zone_Ecoli",
                "Salmonella spp.": "Zone_Salmonella",
            })
        )
        rsm_data = phys_means.merge(anti_mean_wide, on=["Treatment", "Temperature_C", "Time_h"], how="left")
        rsm_data["Zone_Mean_All"] = rsm_data[["Zone_Staph", "Zone_Ecoli", "Zone_Salmonella"]].mean(axis=1)

        rsm_responses = ["WAC", "Aw", "MC_percent", "WSI", "Zone_Mean_All"]
        rsm_models = {response: fit_rsm(rsm_data, response) for response in rsm_responses}
        {response: round(model.rsquared, 4) for response, model in rsm_models.items()}
        """
    ),
    code(
        """
        temp_grid = np.linspace(50, 70, 61)
        time_grid = np.linspace(4, 8, 61)
        TT, HH = np.meshgrid(temp_grid, time_grid)
        pred_grid = pd.DataFrame({
            "Temperature_C": TT.ravel(),
            "Time_h": HH.ravel(),
        })
        pred_grid["Temp2"] = pred_grid["Temperature_C"] ** 2
        pred_grid["Time2"] = pred_grid["Time_h"] ** 2
        pred_grid["Temp_Time"] = pred_grid["Temperature_C"] * pred_grid["Time_h"]

        fig, axes = plt.subplots(1, len(rsm_responses), figsize=(18, 3.8), constrained_layout=True)
        for ax, response in zip(axes, rsm_responses):
            pred = rsm_models[response].predict(pred_grid).to_numpy().reshape(TT.shape)
            contour = ax.contourf(TT, HH, pred, levels=18, cmap="viridis")
            ax.scatter(rsm_data["Temperature_C"], rsm_data["Time_h"], c="white", edgecolor="black", s=35)
            ax.set_title(response)
            ax.set_xlabel("Temperature (°C)")
            ax.set_ylabel("Time (h)")
            fig.colorbar(contour, ax=ax, shrink=.85)
        rsm_contour_path = FIG_DIR / "rsm_contour_surfaces.png"
        fig.savefig(rsm_contour_path, bbox_inches="tight")
        plt.show()
        rsm_contour_path
        """
    ),
    md("## 8. Multi-Response Optimization Screening"),
    code(
        """
        def desirability_max(x: pd.Series) -> pd.Series:
            lo, hi = x.min(), x.max()
            return (x - lo) / (hi - lo) if hi != lo else pd.Series(1, index=x.index)

        def desirability_min(x: pd.Series) -> pd.Series:
            return 1 - desirability_max(x)

        opt = pred_grid[["Temperature_C", "Time_h"]].copy()
        for response, model in rsm_models.items():
            opt[response] = model.predict(pred_grid)

        # Reasonable objective directions from the product goals:
        # lower Aw and moisture for stability; higher WAC, WSI, and antimicrobial zone for functionality.
        opt["d_WAC"] = desirability_max(opt["WAC"])
        opt["d_Aw"] = desirability_min(opt["Aw"])
        opt["d_MC"] = desirability_min(opt["MC_percent"])
        opt["d_WSI"] = desirability_max(opt["WSI"])
        opt["d_Zone"] = desirability_max(opt["Zone_Mean_All"])
        d_cols = ["d_WAC", "d_Aw", "d_MC", "d_WSI", "d_Zone"]
        opt["Composite_Desirability"] = opt[d_cols].prod(axis=1) ** (1 / len(d_cols))

        top_opt = opt.sort_values("Composite_Desirability", ascending=False).head(15)
        top_opt.to_csv(OUT_DIR / "rsm_optimization_top_candidates.csv", index=False)
        top_opt.round(4)
        """
    ),
    code(
        """
        fig, ax = plt.subplots(figsize=(6.5, 5), constrained_layout=True)
        desir = opt["Composite_Desirability"].to_numpy().reshape(TT.shape)
        contour = ax.contourf(TT, HH, desir, levels=18, cmap="mako")
        best = top_opt.iloc[0]
        ax.scatter(best["Temperature_C"], best["Time_h"], s=90, c="red", edgecolor="white", label="Best screened point")
        ax.scatter(rsm_data["Temperature_C"], rsm_data["Time_h"], c="white", edgecolor="black", s=35, label="Observed treatments")
        ax.set_title("Composite Desirability Surface")
        ax.set_xlabel("Temperature (°C)")
        ax.set_ylabel("Time (h)")
        ax.legend(loc="best")
        fig.colorbar(contour, ax=ax, label="Desirability")
        desir_path = FIG_DIR / "rsm_composite_desirability.png"
        fig.savefig(desir_path, bbox_inches="tight")
        plt.show()
        desir_path
        """
    ),
    md("## 9. Key Findings Generated from the Computation"),
    code(
        """
        best_observed = rsm_data.copy()
        best_observed["Observed_Desirability"] = (
            desirability_max(best_observed["WAC"])
            * desirability_min(best_observed["Aw"])
            * desirability_min(best_observed["MC_percent"])
            * desirability_max(best_observed["WSI"])
            * desirability_max(best_observed["Zone_Mean_All"])
        ) ** (1/5)
        best_observed = best_observed.sort_values("Observed_Desirability", ascending=False)

        summary_lines = [
            "Treatment 9 is analyzed as 70°C/8 h to complete the documented 3x3 factorial design.",
            f"Best observed treatment by composite desirability: Treatment {int(best_observed.iloc[0]['Treatment'])} "
            f"({best_observed.iloc[0]['Temperature_C']:.0f}°C, {best_observed.iloc[0]['Time_h']:.0f} h), "
            f"desirability={best_observed.iloc[0]['Observed_Desirability']:.3f}.",
            f"Best RSM-screened condition: {top_opt.iloc[0]['Temperature_C']:.2f}°C and {top_opt.iloc[0]['Time_h']:.2f} h, "
            f"desirability={top_opt.iloc[0]['Composite_Desirability']:.3f}.",
            "Two-way ANOVA and one-way antimicrobial ANOVA result tables were exported to the outputs folder.",
        ]
        for line in summary_lines:
            print(line)

        pd.DataFrame({"Finding": summary_lines}).to_csv(OUT_DIR / "analysis_key_findings.csv", index=False)
        best_observed[["Treatment", "Temperature_C", "Time_h", "WAC", "Aw", "MC_percent", "WSI", "Zone_Mean_All", "Observed_Desirability"]].round(4)
        """
    ),
]

nb["cells"] = cells
nbf.write(nb, NOTEBOOK)
print(NOTEBOOK)
