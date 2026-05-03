# Contributing

Thanks for taking the time to improve Stock Data Manager.

## Local Setup

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

For editable installation:

```bash
pip install -e .
```

## Checks

Before opening a pull request, run:

```bash
python -m py_compile Stock_Data_Manager.py
python Stock_Data_Manager.py --help
```

If you have dependencies installed and internet access, also test a real export:

```bash
python Stock_Data_Manager.py --stock TSLA --period 1mo --interval 1d --no-chart --no-auto-update
```

## Pull Request Notes

- Keep changes focused.
- Do not commit generated Excel files from `stocks/` or `reports/`.
- Include clear usage examples when adding or changing command-line options.
