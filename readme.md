

# Požadavky

Python 3.10+ (možná bude fungovat na nižším)

Knihovny: tkinter, openpyxl, csv

Tento návod je určen pro Windows

## Instalace knihoven

### Instalace globálně pro Python

Knihovny se buď dají instalovat do vlastní instalace Pythonu, a sice přes pip (Pythonovský správce knihoven)

```bash
pip install tkinter openpyxl
```

Pokud pip není uveden v cestě, pak také takto:

```bash
python -m pip install tk openpyxl
```

Knihovny lze také naistalovat z přiloženého requirements.txt

```bash
pip install -r requirements.txt
```

### Instalace do virtuálního prostředí

Knihovny lze také nainstalovat do virtuálního prostředí, které je vytvořeno pro daný projekt. V tomto případě je nutné vytvořit virtuální prostředí a nainstalovat knihovny do něj.

Následující příkazy vytvoří virtuální prostředí, spustí ho a nainstalují knihovny do něj.

```bash
python -m venv moje_virtualni_prostredi
moje_virtualni_prostredi\Scripts\activate.bat
pip install -r requirements.txt
```

# Spuštění

## Spuštění z příkazové řádky

```bash
python src\gui.py
```

Lze spustit také pouze skript, který sloučí soubory

```bash
python src\merge_files.py 1 out.csv in1.xlsx in2.xlsx
```



