# Room Occupancy Estimation — Student Package

**Prepared by Engr. Jamie Eduardo Rosal, MSCpE**

## Start here

1. If you received `Student.zip`, extract the entire archive first.
2. Open **Room_Occupancy_Student_Guide.pdf** and follow the 12 steps.
3. Set up Python using the instructions below.
4. Open **Room_Occupancy_Estimation.ipynb** in JupyterLab and follow the matching numbered sections.

Keep this folder's structure intact. The notebook expects the dataset alongside it and writes its outputs into `results/`.

## Main files

| File | Purpose |
|---|---|
| `Room_Occupancy_Student_Guide.pdf` | Printable step-by-step student guide |
| `Room_Occupancy_Student_Guide.docx` | Editable guide for Word or import into Google Docs |
| `Room_Occupancy_Estimation.ipynb` | Executed Jupyter notebook with code and results |
| `Room_Occupancy_Estimation.html` | View the notebook in a browser without installing Python |
| `Occupancy_Estimation.csv` | Dataset used in the exercises |
| `Intro-to-AI-Proposal.docx` | Original project objectives and proposal |
| `requirements.txt` | Tested direct package versions |
| `build_materials.py` | Rebuild the notebook and both guides |
| `verify_materials.py` | Independently check saved results and documents |
| `REVIEW.md` | Review findings and interpretation limits |
| `results/` | Reference scores, predictions, figures, split assignments and verification records |

## Install and launch

Use **Python 3.12**. Open a terminal in this extracted `Student` folder.

On macOS/Linux:

```bash
python3.12 -m venv .venv
source .venv/bin/activate
```

On Windows Command Prompt:

```bat
py -3.12 -m venv .venv
.venv\Scripts\activate.bat
```

Then run:

```bash
python -m pip install -r requirements.txt
python -m ipykernel install --user --name occupancy-guide --display-name "Occupancy Guide"
python -m jupyterlab
```

Open the notebook, select the **Occupancy Guide** kernel, then choose **Restart Kernel and Run All Cells**. Package installation needs internet; the experiment uses the included CSV offline.

## Rebuild and verify

Run these commands from this folder with the environment active:

```bash
python build_materials.py --verify
python verify_materials.py
```

The builder executes twice, checks repeatability and regenerates the notebook, HTML, Word guide and PDF guide. The verifier independently checks predictions, metrics, split assignments, confusion matrices and document score tables.

To regenerate only the guides from existing results:

```bash
python build_materials.py --documents-only
```

These commands overwrite generated files. Save a separate copy of your own edits first. If you change the builder or data, use the full build instead of documents-only. Notebook execution clears earlier verification records; rerun the full build with `--verify` and the verifier to obtain current records.

## Interpret the results carefully

The original random-split baseline accuracy is **99.70%**. December-to-January evaluation gives **88.02%**, with no class-1 examples in January. Randomly mixed neighboring readings can produce optimistic results. Sensor rankings depend on the evaluation period; a small improvement after removing a group does not prove that group is unnecessary or noisy.

## Troubleshooting

- **Missing CSV:** start Jupyter in this folder and keep `Occupancy_Estimation.csv` beside the notebook.
- **Missing package:** install into the same environment used by the selected kernel.
- **PDF generation:** use the included builder; it does not require LaTeX or a browser.
- **Schema/timestamp assertion:** inspect the data rather than disabling validation.

Your submission should include the executed notebook, requested guide/report, supporting results and a conclusion explaining the evaluation limitations. Follow any additional instructions from your instructor.
