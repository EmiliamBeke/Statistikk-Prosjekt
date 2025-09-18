#!/usr/bin/env python3
# -*- coding: utf-8 -*-
#test
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pathlib import Path

# 1) Pek på Excel-filen (endre hvis nødvendig)
excel_path = Path(r"C:\Users\isakr\Documents\statistikk\sko_str-høyde.xlsx")
assert excel_path.exists(), f"Finner ikke {excel_path.resolve()}"

# 2) Les første ark
xls = pd.ExcelFile(excel_path)
print("Ark i arbeidsboken:", xls.sheet_names)
df = pd.read_excel(excel_path, sheet_name=xls.sheet_names[0])

# 3) Inspeksjon
print(df.head(10))
df.info()

# 4) Standardiser kolonnenavn
df.columns = (df.columns
              .str.strip()
              .str.replace(r"\s+", "_", regex=True)
              .str.replace(r"[^\w_]", "", regex=True)
              .str.lower())

# Håndter ev. norske bokstaver i navn (valgfritt, gjør koden robust)
if "høyde" in df.columns:
    df = df.rename(columns={"høyde": "hoyde"})
if "skostørrelse" in df.columns:
    df = df.rename(columns={"skostørrelse": "skostorrelse"})

# --- Regresjon: y = a + b*x (x = skostorrelse, y = hoyde) ---
df["hoyde"] = pd.to_numeric(df["hoyde"], errors="coerce")
df["skostorrelse"] = pd.to_numeric(df["skostorrelse"], errors="coerce")
reg_df = df[["skostorrelse", "hoyde"]].dropna()

x = reg_df["skostorrelse"].to_numpy()
y = reg_df["hoyde"].to_numpy()
if x.size < 2:
    raise ValueError("For få datapunkter til å gjøre lineær regresjon.")

# polyfit(grad=1) -> [stigning b, konstant a]
b, a = np.polyfit(x, y, 1)          # y ≈ a + b*x
y_hat = a + b*x

# R^2
ss_res = np.sum((y - y_hat)**2)
ss_tot = np.sum((y - np.mean(y))**2)
r2 = 1 - ss_res/ss_tot if ss_tot > 0 else float("nan")

print("\nRegresjonslinje: y = a + b*x")
print(f"a (konstantledd) = {a:.3f}")
print(f"b (stigning)     = {b:.3f}")
print(f"R^2              = {r2:.3f}")

# 5) Plot + lagre PDF
x_line = np.linspace(x.min(), x.max(), 200)
y_line = a + b * x_line

plt.figure(figsize=(6, 4))
plt.scatter(x, y, label="Data")
plt.plot(x_line, y_line, label=f"y = {a:.2f} + {b:.2f}x")
plt.xlabel("Skostørrelse")
plt.ylabel("Høyde (cm)")
plt.title("Høyde som funksjon av skostørrelse")
plt.legend()
plt.tight_layout()

out_dir = Path("utdata")
out_dir.mkdir(exist_ok=True)
pdf_path = out_dir / "regresjon_sko_vs_hoyde.pdf"
plt.savefig(pdf_path)
plt.show()

# 6) Lagre parametere (+ valgfritt: renset rådata)
with open(out_dir / "regresjon_parametre.txt", "w", encoding="utf-8") as f:
    f.write("Regresjon: y = a + b*x (x=skostorrelse, y=hoyde)\n")
    f.write(f"a (konstantledd): {a:.6f}\n")
    f.write(f"b (stigning):     {b:.6f}\n")
    f.write(f"R^2:              {r2:.6f}\n")
    f.write(f"Antall punkt:     {len(reg_df)}\n")

df.to_excel(out_dir / "raadata.xlsx", index=False)  # valgfritt

print("Figur lagret til:", pdf_path.resolve())
print("Parametere lagret til:", (out_dir / "regresjon_parametre.txt").resolve())
