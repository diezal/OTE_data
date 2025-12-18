import pandas as pd

#Načtení DT ČR

# Relativní cesta k souboru (podsložka Data)
excel_path = r"Data/Rocni_zprava_2025_V0_trhy_ERD.xlsx"
sheet_name = "DT ČR"

# Načtení celého listu bez hlaviček
DT_CR = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)

# Výřez buněk A6:M12417
df_DT_CR = DT_CR.iloc[5:12417, 0:13]

# První řádek výřezu jako hlavička
df_DT_CR.columns = df_DT_CR.iloc[0]
df_DT_CR = df_DT_CR[1:].reset_index(drop=True)

print(df_DT_CR.head())
df_DT_CR = df_DT_CR.dropna(subset=[df_DT_CR.columns[3]])




#Načtení Indexy Indexy DT

sheet_name = "Indexy ČR"
Indexy_DT = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)

# Výřez buněk A6:M12417
df_Indexy_DT = Indexy_DT.iloc[5:339, 0:7]

# První řádek výřezu jako hlavička
df_Indexy_DT.columns = df_Indexy_DT.iloc[0]
df_Indexy_DT = df_Indexy_DT[1:].reset_index(drop=True)

print(df_Indexy_DT.head())
df_Indexy_DT = df_Indexy_DT.dropna(subset=[df_Indexy_DT.columns[3]])



#Načtení přeshraničních toků

sheet_name = "DT ČR Import-Export"
DT_CR_Import_Export = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)

# Výřez buněk A6:M12417
df_DT_CR_Import_Export = DT_CR_Import_Export.iloc[5:12417, 0:11]

# První řádek výřezu jako hlavička
df_DT_CR_Import_Export.columns = df_DT_CR_Import_Export.iloc[0]
df_DT_CR_Import_Export = df_DT_CR_Import_Export[1:].reset_index(drop=True)

print(df_DT_CR_Import_Export.head())
df_DT_CR_Import_Export = df_DT_CR_Import_Export.dropna(subset=[df_DT_CR_Import_Export.columns[3]])


df_DT_CR_Import_Export.columns = (
    ["Den", "Perioda", "Časový interval"]
    + df_DT_CR_Import_Export.columns[3:].tolist()
)



#Načtení VDT (EUR)
sheet_name = "VDT (EUR)"

# Načtení celého listu bez hlaviček
VDT = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)

# Výřez buněk A6:M12417
df_VDT = VDT.iloc[5:32070, 0:15]

# První řádek výřezu jako hlavička
df_VDT.columns = df_VDT.iloc[0]
df_VDT = df_VDT[1:].reset_index(drop=True)

print(df_VDT.head())
df_VDT = df_VDT.dropna(subset=[df_VDT.columns[3]])



#Načtení IDA ČR

sheet_name = "IDA ČR"
IDA_CR = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)

# Výřez buněk A6:AG32070 (33 sloupců)
df_IDA_CR = IDA_CR.iloc[5:32070, 0:33]

# Nastavit první řádek jako hlavičku
df_IDA_CR.columns = df_IDA_CR.iloc[0]
df_IDA_CR = df_IDA_CR[1:].reset_index(drop=True)

# ❗ Odstranit řádky, kde je NaN pouze ve 4. sloupci (index 3)
df_IDA_CR = df_IDA_CR.dropna(subset=[df_IDA_CR.columns[3]])

# ✔ Přejmenovat první tři sloupce správně
df_IDA_CR.columns = (
    ["Den", "Perioda", "Časový interval"]
    + df_IDA_CR.columns[3:].tolist()
)

print(df_IDA_CR.head())

output = r"C:\Users\David Nezval\Desktop\IDA_CR_clean.xlsx"
IDA_CR.to_excel(output, index=False)



#Načtení IDA ČR

# načti list s reálnou hlavičkou na řádku 5
df = pd.read_excel(excel_path, sheet_name="IDA ČR", header=4)

# odstranění prázdných řádků v periodě
df = df.dropna(subset=["Perioda"]).reset_index(drop=True)

# základní sloupce
base_cols = ["Den", "Perioda", "Časový interval"]

# indexy začátků bloků
ida1_start = 3
ida2_start = 12
ida3_start = 21

# názvy sloupců, které chceme vždy
measurement_cols = [
    "Nákup (MWh)",
    "Prodej (MWh)",
    "Saldo IDA\n(MWh)",
    "Export\n(MWh)",
    "Import\n(MWh)",
    "Množství - vč. Exp a Imp (MWh)",
    "Marginální cena ČR (EUR/MWh)",
    "Marginální cena ČR (Kč/MWh)",
    "Kurz Kč/EUR (ČNB)"
]

# vytvoření IDA1
IDA1 = df[base_cols + df.columns[ida1_start:ida1_start + len(measurement_cols)].tolist()].copy()
IDA1.columns = base_cols + measurement_cols

# vytvoření IDA2
IDA2 = df[base_cols + df.columns[ida2_start:ida2_start + len(measurement_cols)].tolist()].copy()
IDA2.columns = base_cols + measurement_cols

# vytvoření IDA3
IDA3 = df[base_cols + df.columns[ida3_start:ida3_start + len(measurement_cols)].tolist()].copy()
IDA3.columns = base_cols + measurement_cols




#Načti Index IDA

excel_path = r"Data/Rocni_zprava_2025_V0_trhy_ERD.xlsx"

# Načtení listu "Indexy IDA" – hlavička je v řádku 5 (index 4)
idx = pd.read_excel(excel_path, sheet_name="Indexy IDA", header=4)

# První řádek po hlavičce obsahuje názvy indexů (BASE/PEAK/OFFPEAK atd.)
header_row = idx.iloc[0]

# Skutečná data začínají od dalšího řádku
data = idx.iloc[1:].reset_index(drop=True)

# ---- INDEX_IDA1 ----
# Sloupce pro IDA 1 jsou od 1 do 6 (včetně)
ida1_cols = idx.columns[1:7]                  # původní názvy ('IDA 1', 'Unnamed: 2', ...)
ida1_names = header_row.iloc[1:7].tolist()    # nové názvy (BASE LOAD..., PEAK..., ...)

INDEX_IDA1 = data[['Den'] + ida1_cols.tolist()].copy()
INDEX_IDA1.columns = ['Den'] + ida1_names     # přejmenujeme sloupce podle řádku s názvy indexů

# ---- INDEX_IDA2 ----
# Sloupce pro IDA 2 jsou od 7 do 12 (včetně)
ida2_cols = idx.columns[7:13]
ida2_names = header_row.iloc[7:13].tolist()

INDEX_IDA2 = data[['Den'] + ida2_cols.tolist()].copy()
INDEX_IDA2.columns = ['Den'] + ida2_names

# (volitelně) převod data na datetime
INDEX_IDA1['Den'] = pd.to_datetime(INDEX_IDA1['Den'])
INDEX_IDA2['Den'] = pd.to_datetime(INDEX_IDA2['Den'])

# Kontrola
print(INDEX_IDA1.head())
print(INDEX_IDA2.head())



#Načtení Odchylky

# Relativní cesta k souboru (podsložka Data)
excel_path = r"Data/Rocni_zprava_2025_V0_odchylky.xlsx"
sheet_name = "Odchylky"

# Načtení celého listu bez hlaviček
odchylky = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)

# Výřez buněk A6:M12417
df_odchylky = odchylky.iloc[5:32070, 0:17]

# První řádek výřezu jako hlavička
df_odchylky.columns = df_odchylky.iloc[0]
df_odchylky = df_odchylky[1:].reset_index(drop=True)

print(df_odchylky.head())
df_odchylky = df_odchylky.dropna(subset=[df_odchylky.columns[3]])