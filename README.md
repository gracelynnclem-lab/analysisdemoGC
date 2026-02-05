# analysisdemoGC

## JE analysis workflow

This repo includes a repeatable, no-dependency analysis workflow for journal entry (JE) exports.

### Run locally

```bash
python analysis/je_basic_analysis.py
```

By default, the script reads `je_samples.xlsx` and writes artifacts to `analysis_outputs/`.
You can override the input and output locations:

```bash
python analysis/je_basic_analysis.py --input path/to/je.xlsx --output-dir path/to/output
```

### GitHub Actions

The `JE Analysis` workflow runs automatically on pushes that touch the Excel file, the analysis
script, or the workflow itself. It uploads `analysis_outputs/` as a build artifact.
