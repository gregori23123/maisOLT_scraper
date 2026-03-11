# MaisOLT Scraper

Coleta estatísticas de ONUs por OLT no painel maisOLT e exporta para Excel.

## Requisitos

- Python 3.8+
- Google Chrome instalado

## Instalação

```bash
pip install selenium webdriver-manager pandas openpyxl
```

## Uso

```bash
python maisolt_scraper.py
```

Informe e-mail e senha quando solicitado. O arquivo `.xlsx` será salvo na mesma pasta.

## Configuração

Edite no início do script:

```python
BASE_URL = "https://suaempresa.maisolt.com.br"
OLT_IDS  = [i for i in range(1, 16) if i not in (7, 10, 14)]
```

## Gerar .exe

```bash
pip install pyinstaller
pyinstaller --onefile --name "MaisOLT_Scraper" maisolt_scraper.py
```
