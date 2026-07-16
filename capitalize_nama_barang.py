import sqlite3

DB_PATH = "nota_belanja.db"

def capitalize_each_word():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    # ===== Proses tabel ItemBelanja =====
    print("🔹 Memproses tabel ItemBelanja...")
    cur.execute("SELECT id, nama_barang FROM ItemBelanja WHERE nama_barang IS NOT NULL AND nama_barang <> ''")
    rows = cur.fetchall()
    updated_belanja = 0

    for rid, nama in rows:
        proper = nama.title().strip()
        if proper != nama:
            cur.execute("UPDATE ItemBelanja SET nama_barang = ? WHERE id = ?", (proper, rid))
            updated_belanja += 1

    # ===== Proses tabel ItemPenjualan =====
    print("🔹 Memproses tabel ItemPenjualan...")
    cur.execute("SELECT id, nama_barang FROM ItemPenjualan WHERE nama_barang IS NOT NULL AND nama_barang <> ''")
    rows = cur.fetchall()
    updated_penjualan = 0

    for rid, nama in rows:
        proper = nama.title().strip()
        if proper != nama:
            cur.execute("UPDATE ItemPenjualan SET nama_barang = ? WHERE id = ?", (proper, rid))
            updated_penjualan += 1

    conn.commit()
    conn.close()

    print(f"✅ Selesai! {updated_belanja} baris di ItemBelanja dan {updated_penjualan} baris di ItemPenjualan diperbarui.")
    print("📁 Database:", DB_PATH)

if __name__ == "__main__":
    capitalize_each_word()
