# Zebra Organiser

PyQt6 GUI aplikace pro tisk štítků na Zebra tiskárně ze souboru Excel.

## Popis

Aplikace načte data z `organiser.xlsx` a umožňuje tisknout štítky na Zebra tiskárnu přes ZPL protokol. Každý řádek v Excelu = jeden štítek.

## Požadavky

- Python 3.8+
- PyQt6
- openpyxl
- zebra

Instalace závislostí:

```bash
pip install PyQt6 openpyxl zebra-day
```

## Spuštění

```bash
python3 organiser.py
```

## Soubor organiser.xlsx

Soubor `organiser.xlsx` obsahuje data pro tisk. Každý řádek odpovídá jednomu štítku — buňky v řádku se vytisknou jako text na štítek (každá buňka na samostatný řádek).

## Funkce

- Výběr tiskárny ze systému (lpstat / Zebra knihovna)
- Nastavení velikosti štítku (50x25mm, 4x6", ...)
- Nastavení DPI (203 / 300)
- Volitelný obdélník kolem textu
- Automatické nahrazení diakritiky pro ZPL kompatibilitu
- Zapamatování poslední použité tiskárny
