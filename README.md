# Reportus

A minimal Windows desktop application for building polished financial reports from clean uploaded data and explicitly requested external research.

## Development

```powershell
py -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Run the non-UI test suite with:

```powershell
python -m unittest discover -s tests -v
```
