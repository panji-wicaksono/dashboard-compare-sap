import sqlite3
conn = sqlite3.connect("dashboard.db")
conn.row_factory = sqlite3.Row

print("=== AUFM (material 1132000W) ===")
rows = conn.execute("SELECT AUFNR, MATNR, BWART, MENGE, BLDAT FROM aufm WHERE MATNR LIKE ?", ("%1132000W%",)).fetchall()
for r in rows:
    print(dict(r))
if not rows:
    print("(tidak ada data)")

print()
print("=== PRODSYS (material 1132000W) ===")
rows = conn.execute("SELECT AUFNR, MATNR, MENGE, GSTRP, AFRUD_ISDD FROM prodsys WHERE MATNR LIKE ?", ("%1132000W%",)).fetchall()
for r in rows:
    print(dict(r))
if not rows:
    print("(tidak ada data)")

print()
print("=== BLDAT unik di AUFM ===")
for r in conn.execute("SELECT DISTINCT BLDAT FROM aufm ORDER BY BLDAT"):
    print(r[0])

print()
print("=== GSTRP unik di Prodsys ===")
for r in conn.execute("SELECT DISTINCT GSTRP FROM prodsys WHERE GSTRP IS NOT NULL ORDER BY GSTRP"):
    print(r[0])

conn.close()
