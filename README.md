# Colab Merge Project

This repo shows a minimal layout for a Colab-friendly tool that merges two tabular files (CSV/XLSX) using a small Python package (`yourpkg`).

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/mctwork2/tpok003/blob/REF/notebooks/colab_merge_tool.ipynb)

## Quick start (locally)

```bash
pip install -U pip
pip install -e .
python -c "from yourpkg.merge import merge_two_files; print(merge_two_files('examples/example_first.csv','examples/example_second.csv', mode='join', how='inner', left_key='id', right_key='id', out_path='merged.xlsx'))"
```

## What’s inside

- `pyproject.toml` — project metadata and dependencies
- `yourpkg/` — importable package with all logic in functions
- `notebooks/colab_merge_tool.ipynb` — a Colab form/demo that calls functions from the package
- `examples/` — small sample files (plus your uploaded files copied here)
- `scripts/` — place for standalone scripts (renamed to `nbutest.py`)

## After you push to GitHub

1. Replace `mctwork2/tpok003/REF` in this README and in the notebook’s install cell.
2. Click the **Open in Colab** badge to run the demo.


## Web UI в Colab
В ноутбуке есть блок **Gradio UI**: загрузите два файла (CSV/XLSX), введите дату (опционально) и получите скачиваемый результат `merged_result_YYYY-MM-DD.xlsx`. По умолчанию объединение идёт по колонке `id` (inner-join).
