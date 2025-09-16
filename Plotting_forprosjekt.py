#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue Sep 16 13:52:58 2025

@author: olehartvigistad
"""

# 0) Installer (kun hvis nødvendig – kjør i én celle)
# !pip install pandas openpyxl

import pandas as pd
from pathlib import Path

# 1) Pek på Excel-filen
excel_path = Path("data.xlsx")  # <- sett riktig navn/sti, f.eks. "/home/jovyan/work/min_fil.xlsx"
assert excel_path.exists(), f"Finner ikke {excel_path.resolve()}"

# 2) Finn tilgjengelige ark
xls = pd.ExcelFile(excel_path)
print("Ark i arbeidsboken:", xls.sheet_names)

# 3) Les et ark (navn eller indeks). Bruk sheet_names[0] for første ark.
sheet_to_read = xls.sheet_names[0]
df = pd.read_excel(excel_path, sheet_name=sheet_to_read)

# 4) Kjapp inspeksjon
display(df.head(10))
print(df.info())

# 5) (Valgfritt) Standardiser kolonnenavn
df.columns = (
    df.columns
      .str.strip()
      .str.replace(r"\s+", "_", regex=True)
      .str.replace(r"[^\w_]", "", regex=True)
      .str.lower()
)
display(df.head(3))

# 6) (Valgfritt) Konverter dato-/tallkolonner
# Bytt ut 'dato' og 'beløp' med reelle kolonnenavn i filen din
for col in ['dato']:
    if col in df.columns:
        df[col] = pd.to_datetime(df[col], errors='coerce')

for col in ['beløp', 'amount', 'sum']:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='coerce')

# 7) Eksempel: filtrering
# Behold rader fra 2024 og nyere (hvis du har en datokolonne)
if 'dato' in df.columns:
    df_2024 = df[df['dato'].dt.year >= 2024]
else:
    df_2024 = df.copy()

# 8) Eksempel: gruppér og summer
# Bytt 'kategori' og 'beløp' til det som gir mening i dine data
if {'kategori', 'beløp'}.issubset(df_2024.columns):
    oppsummert = (
        df_2024
        .groupby('kategori', dropna=False, as_index=False)['beløp']
        .sum()
        .sort_values('beløp', ascending=False)
    )
else:
    oppsummert = df_2024.head(0)  # tomt rammeverk hvis kolonner mangler

display(oppsummert)

# 9) (Valgfritt) Pivot-tabell (endre kolonner etter behov)
# Eksempel: summer 'beløp' per 'kategori' per 'måned'
if {'kategori', 'beløp', 'dato'}.issubset(df.columns):
    df['måned'] = df['dato'].dt.to_period('M').dt.to_timestamp()
    pivot = pd.pivot_table(
        df,
        values='beløp',
        index='kategori',
        columns='måned',
        aggfunc='sum',
        fill_value=0,
        margins=True
    )
    display(pivot)

# 10) (Valgfritt) En enkel graf
import matplotlib.pyplot as plt

if not oppsummert.empty and 'kategori' in oppsummert and 'beløp' in oppsummert:
    ax = oppsummert.plot(kind='bar', x='kategori', y='beløp', legend=False, figsize=(8,4))
    ax.set_title('Sum per kategori')
    ax.set_xlabel('Kategori')
    ax.set_ylabel('Beløp')
    plt.tight_layout()
    plt.show()

# 11) Lagre resultater
out_dir = Path("utdata")
out_dir.mkdir(exist_ok=True)

# til CSV
df_2024.to_csv(out_dir / "filtrert_2024.csv", index=False)
oppsummert.to_csv(out_dir / "oppsummering_per_kategori.csv", index=False)

# til Excel med flere ark
with pd.ExcelWriter(out_dir / "resultater.xlsx", engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="rådata", index=False)
    df_2024.to_excel(writer, sheet_name="filtrert_2024", index=False)
    if not oppsummert.empty:
        oppsummert.to_excel(writer, sheet_name="oppsummering", index=False)
    try:
        pivot.to_excel(writer, sheet_name="pivot")
    except NameError:
        pass

print("Ferdig! Filer skrevet til:", out_dir.resolve())