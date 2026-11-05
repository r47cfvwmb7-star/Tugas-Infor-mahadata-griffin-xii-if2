"""Praktikum Mahadata: Analisis Pola Penyewaan dan Operasional Pusat Olahraga."""
from pathlib import Path
import sqlite3
import pandas as pd
import matplotlib.pyplot as plt

FOLDER = Path(__file__).resolve().parent
DB_FILE = FOLDER / "database.db"
GRAFIK_DIR = FOLDER / "grafik"
GRAFIK_DIR.mkdir(exist_ok=True)
EXCEL_FILE = FOLDER / "hasil_export.xlsx"
CSV_FILE = FOLDER / "hasil_export.csv"

NAMA_HARI = {
    "Monday": "Senin", "Tuesday": "Selasa", "Wednesday": "Rabu",
    "Thursday": "Kamis", "Friday": "Jumat", "Saturday": "Sabtu",
    "Sunday": "Minggu"
}

def baca_database(conn):
    """Membaca tabel database ke DataFrame menggunakan Pandas."""
    customers = pd.read_sql_query("SELECT * FROM customers", conn)
    sports = pd.read_sql_query("SELECT * FROM sports", conn)
    courts = pd.read_sql_query("SELECT * FROM courts", conn)
    bookings = pd.read_sql_query("SELECT * FROM bookings", conn)
    details = pd.read_sql_query("SELECT * FROM booking_details", conn)
    payments = pd.read_sql_query("SELECT * FROM payments", conn)
    reviews = pd.read_sql_query("SELECT * FROM reviews", conn)

    # JOIN menghasilkan satu baris untuk setiap detail penyewaan lapangan.
    detail_query = """
    SELECT
        bd.booking_detail_id,
        b.booking_id,
        b.booking_date,
        b.booking_status,
        cu.customer_id,
        cu.customer_name,
        cu.district,
        sp.sport_name,
        co.court_name,
        co.court_type,
        bd.start_time,
        bd.duration_hours,
        bd.hourly_rate,
        bd.subtotal,
        p.payment_method,
        p.amount_paid,
        p.payment_status,
        r.rating
    FROM booking_details AS bd
    JOIN bookings AS b ON b.booking_id = bd.booking_id
    JOIN customers AS cu ON cu.customer_id = b.customer_id
    JOIN courts AS co ON co.court_id = bd.court_id
    JOIN sports AS sp ON sp.sport_id = co.sport_id
    LEFT JOIN payments AS p ON p.booking_id = b.booking_id
    LEFT JOIN reviews AS r
        ON r.booking_id = b.booking_id AND r.court_id = bd.court_id
    ORDER BY b.booking_date, b.booking_id, bd.booking_detail_id
    """
    detail_data = pd.read_sql_query(detail_query, conn)
    return {
        "customers": customers, "sports": sports, "courts": courts,
        "bookings": bookings, "booking_details": details,
        "payments": payments, "reviews": reviews,
        "detail_data": detail_data
    }

def bersihkan_data(tables):
    """Membersihkan format teks, tanggal, angka, dan baris duplikat."""
    data = tables["detail_data"].copy()

    kolom_teks = [
        "booking_status", "customer_name", "district", "sport_name",
        "court_name", "court_type", "start_time", "payment_method",
        "payment_status"
    ]
    for kolom in kolom_teks:
        data[kolom] = data[kolom].astype("string").str.strip()
        data[kolom] = data[kolom].replace("", pd.NA)

    data["booking_date"] = pd.to_datetime(data["booking_date"], errors="coerce")
    kolom_angka = [
        "booking_detail_id", "booking_id", "customer_id", "duration_hours",
        "hourly_rate", "subtotal", "amount_paid", "rating"
    ]
    for kolom in kolom_angka:
        data[kolom] = pd.to_numeric(data[kolom], errors="coerce")

    jumlah_awal = len(data)
    data = data.drop_duplicates(subset=["booking_detail_id"])
    # Hapus baris yang tidak memiliki identitas atau nilai inti yang valid.
    data = data.dropna(subset=[
        "booking_detail_id", "booking_id", "booking_date",
        "customer_id", "court_name", "duration_hours", "subtotal"
    ])
    data = data[(data["duration_hours"] > 0) & (data["subtotal"] >= 0)].copy()
    data["hari"] = data["booking_date"].dt.day_name().map(NAMA_HARI)
    data["bulan"] = data["booking_date"].dt.strftime("%Y-%m")
    data["jam"] = data["start_time"].str.slice(0, 2) + ":00"

    audit = pd.DataFrame([
        ["Baris sebelum pembersihan", jumlah_awal],
        ["Baris setelah pembersihan", len(data)],
        ["Duplikat yang dihapus", jumlah_awal - len(data)],
        ["Kolom wajib yang masih kosong", int(data[
            ["booking_detail_id", "booking_id", "booking_date",
             "customer_id", "court_name", "duration_hours", "subtotal"]
        ].isna().sum().sum())]
    ], columns=["Pemeriksaan", "Jumlah"])
    return data, audit

def analisis_data(tables, data):
    """Mengolah DataFrame menjadi ringkasan yang menjawab pertanyaan bisnis."""
    bookings = tables["bookings"].copy()
    payments = tables["payments"].copy()
    reviews = tables["reviews"].copy()

    bookings["booking_date"] = pd.to_datetime(bookings["booking_date"], errors="coerce")
    bookings["bulan"] = bookings["booking_date"].dt.strftime("%Y-%m")
    bookings["hari"] = bookings["booking_date"].dt.day_name().map(NAMA_HARI)

    # Status booking: melihat proporsi transaksi selesai, tertunda, dan batal.
    status = (bookings.groupby("booking_status").size()
              .rename("jumlah_booking").reset_index()
              .rename(columns={"booking_status": "status"}))
    status["persentase"] = (status["jumlah_booking"] / len(bookings) * 100).round(1)

    # Gabungkan detail dengan status pembayaran; satu detail = satu sesi lapangan.
    paid_details = data[
        (data["booking_status"] == "Completed") &
        (data["payment_status"] == "Paid")
    ].copy()

    sport_performance = (paid_details.groupby("sport_name")
        .agg(jumlah_booking=("booking_id", "nunique"),
             jumlah_sesi=("booking_detail_id", "count"),
             total_jam=("duration_hours", "sum"),
             pendapatan=("subtotal", "sum"))
        .reset_index()
        .sort_values("pendapatan", ascending=False)
        .rename(columns={"sport_name": "cabang_olahraga"}))

    # Jam ramai dihitung per detail lapangan, bukan per transaksi gabungan.
    busy_hours = (data[data["booking_status"] == "Completed"]
        .groupby("jam")
        .agg(jumlah_sesi=("booking_detail_id", "count"),
             total_jam=("duration_hours", "sum"))
        .reset_index()
        .sort_values("jam"))

    weekday_demand = (bookings[bookings["booking_status"] == "Completed"]
        .groupby("hari").size().rename("jumlah_booking").reset_index())
    urutan_hari = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
    weekday_demand["hari"] = pd.Categorical(
        weekday_demand["hari"], categories=urutan_hari, ordered=True
    )
    weekday_demand = weekday_demand.sort_values("hari")
    weekday_demand["hari"] = weekday_demand["hari"].astype(str)

    paid_bookings = bookings.merge(payments, on="booking_id", how="inner")
    paid_bookings = paid_bookings[
        (paid_bookings["booking_status"] == "Completed") &
        (paid_bookings["payment_status"] == "Paid")
    ].copy()
    monthly_revenue = (paid_bookings.groupby("bulan")["amount_paid"]
        .sum().rename("pendapatan").reset_index().sort_values("bulan"))

    payment_methods = (payments[payments["payment_status"] == "Paid"]
        .groupby("payment_method")
        .agg(jumlah_transaksi=("payment_id", "count"),
             total_pembayaran=("amount_paid", "sum"))
        .reset_index()
        .sort_values("jumlah_transaksi", ascending=False)
        .rename(columns={"payment_method": "metode_pembayaran"}))

    completed_counts = (bookings[bookings["booking_status"] == "Completed"]
        .groupby("customer_id").size())
    customer_counts = tables["customers"][["customer_id"]].copy()
    customer_counts["booking_selesai"] = (
        customer_counts["customer_id"].map(completed_counts).fillna(0).astype(int)
    )
    def kelompokkan(n):
        if n == 0: return "Belum pernah selesai"
        if n == 1: return "Satu kali"
        if n <= 4: return "2–4 kali"
        return "5 kali atau lebih"
    customer_counts["segmen"] = customer_counts["booking_selesai"].map(kelompokkan)
    urutan_segmen = ["Belum pernah selesai", "Satu kali", "2–4 kali", "5 kali atau lebih"]
    customer_segments = (customer_counts.groupby("segmen").size()
        .reindex(urutan_segmen, fill_value=0).rename("jumlah_pelanggan")
        .reset_index())

    court_type = (paid_details.groupby("court_type")
        .agg(jumlah_sesi=("booking_detail_id", "count"),
             total_jam=("duration_hours", "sum"),
             pendapatan=("subtotal", "sum"))
        .reset_index()
        .rename(columns={"court_type": "tipe_lapangan"})
        .sort_values("pendapatan", ascending=False))

    ratings = reviews.merge(tables["courts"][["court_id", "sport_id"]], on="court_id")
    ratings = ratings.merge(tables["sports"][["sport_id", "sport_name"]], on="sport_id")
    ratings_by_sport = (ratings.groupby("sport_name")
        .agg(jumlah_ulasan=("review_id", "count"),
             rata_rating=("rating", "mean"))
        .reset_index()
        .sort_values("rata_rating", ascending=False)
        .rename(columns={"sport_name": "cabang_olahraga"}))
    ratings_by_sport["rata_rating"] = ratings_by_sport["rata_rating"].round(2)

    cancellation = bookings.groupby("bulan").agg(
        seluruh_booking=("booking_id", "count"),
        dibatalkan=("booking_status", lambda s: (s == "Cancelled").sum())
    ).reset_index()
    cancellation["tingkat_pembatalan"] = (
        cancellation["dibatalkan"] / cancellation["seluruh_booking"] * 100
    ).round(1)

    return {
        "status_booking": status,
        "kinerja_olahraga": sport_performance,
        "jam_ramai": busy_hours,
        "hari_ramai": weekday_demand,
        "pendapatan_bulanan": monthly_revenue,
        "metode_pembayaran": payment_methods,
        "segmen_pelanggan": customer_segments,
        "tipe_lapangan": court_type,
        "rating_olahraga": ratings_by_sport,
        "pembatalan_bulanan": cancellation
    }

def buat_grafik(hasil):
    """Menyimpan grafik dengan judul dan label sumbu yang jelas."""
    # 1. Pola jam ramai
    df = hasil["jam_ramai"]
    plt.figure(figsize=(9, 5))
    plt.bar(df["jam"], df["jumlah_sesi"])
    plt.title("Pola Penyewaan Lapangan Berdasarkan Jam")
    plt.xlabel("Jam mulai penyewaan")
    plt.ylabel("Jumlah sesi selesai")
    plt.grid(axis="y", alpha=0.25)
    plt.tight_layout()
    plt.savefig(GRAFIK_DIR / "01_pola_jam_penyewaan.png", dpi=160)
    plt.close()

    # 2. Perubahan pendapatan bulanan
    df = hasil["pendapatan_bulanan"]
    plt.figure(figsize=(10, 5))
    plt.plot(df["bulan"], df["pendapatan"], marker="o", linewidth=2)
    plt.title("Perubahan Pendapatan Penyewaan per Bulan")
    plt.xlabel("Bulan (2025)")
    plt.ylabel("Pendapatan (Rupiah)")
    plt.xticks(rotation=45, ha="right")
    plt.grid(alpha=0.25)
    plt.tight_layout()
    plt.savefig(GRAFIK_DIR / "02_tren_pendapatan_bulanan.png", dpi=160)
    plt.close()

    # 3. Perbandingan minat cabang olahraga
    df = hasil["kinerja_olahraga"].sort_values("jumlah_booking")
    plt.figure(figsize=(8, 5))
    plt.barh(df["cabang_olahraga"], df["jumlah_booking"])
    plt.title("Permintaan Penyewaan Menurut Cabang Olahraga")
    plt.xlabel("Jumlah booking selesai dan dibayar")
    plt.ylabel("Cabang olahraga")
    plt.tight_layout()
    plt.savefig(GRAFIK_DIR / "03_permintaan_cabang_olahraga.png", dpi=160)
    plt.close()

    # 4. Komposisi status booking
    df = hasil["status_booking"]
    plt.figure(figsize=(7, 5))
    plt.pie(df["jumlah_booking"], labels=df["status"],
            autopct="%1.1f%%", startangle=90)
    plt.title("Komposisi Status Booking")
    plt.tight_layout()
    plt.savefig(GRAFIK_DIR / "04_komposisi_status_booking.png", dpi=160)
    plt.close()

    # 5. Rata-rata rating per cabang olahraga
    df = hasil["rating_olahraga"].sort_values("rata_rating")
    plt.figure(figsize=(8, 5))
    plt.barh(df["cabang_olahraga"], df["rata_rating"])
    plt.xlim(0, 5)
    plt.title("Rata-rata Rating Menurut Cabang Olahraga")
    plt.xlabel("Rata-rata rating (skala 1–5)")
    plt.ylabel("Cabang olahraga")
    plt.tight_layout()
    plt.savefig(GRAFIK_DIR / "05_rating_per_olahraga.png", dpi=160)
    plt.close()

    # 6. Perubahan tingkat pembatalan bulanan
    df = hasil["pembatalan_bulanan"]
    plt.figure(figsize=(10, 5))
    plt.plot(df["bulan"], df["tingkat_pembatalan"], marker="o")
    plt.title("Perubahan Tingkat Pembatalan Booking")
    plt.xlabel("Bulan (2025)")
    plt.ylabel("Booking dibatalkan (%)")
    plt.xticks(rotation=45, ha="right")
    plt.grid(alpha=0.25)
    plt.tight_layout()
    plt.savefig(GRAFIK_DIR / "06_tren_pembatalan.png", dpi=160)
    plt.close()

def buat_ringkasan(tables, hasil, audit):
    booking = tables["bookings"]
    total_booking = len(booking)
    selesai = int((booking["booking_status"] == "Completed").sum())
    batal = int((booking["booking_status"] == "Cancelled").sum())
    pendapatan = int(tables["payments"].loc[
        tables["payments"]["payment_status"] == "Paid", "amount_paid"
    ].sum())
    sport_teratas = hasil["kinerja_olahraga"].sort_values(
        "jumlah_booking", ascending=False
    ).iloc[0]
    jam_teramai = hasil["jam_ramai"].sort_values(
        "jumlah_sesi", ascending=False
    ).iloc[0]
    bulan_tertinggi = hasil["pendapatan_bulanan"].sort_values(
        "pendapatan", ascending=False
    ).iloc[0]
    ringkasan = pd.DataFrame([
        ["Jumlah booking", total_booking],
        ["Booking selesai", selesai],
        ["Booking dibatalkan", batal],
        ["Tingkat pembatalan", round(batal / total_booking * 100, 1)],
        ["Total pembayaran berstatus Paid (Rp)", pendapatan],
        ["Cabang dengan booking terbanyak", sport_teratas["cabang_olahraga"]],
        ["Jam mulai paling ramai", jam_teramai["jam"]],
        ["Bulan dengan pendapatan tertinggi", bulan_tertinggi["bulan"]]
    ], columns=["Indikator", "Hasil"])
    return ringkasan

def main():
    if not DB_FILE.exists():
        raise FileNotFoundError(
            f"Database tidak ditemukan: {DB_FILE}. Letakkan database.db satu folder dengan analisis_data.py."
        )

    conn = sqlite3.connect(DB_FILE)
    try:
        print("=" * 68)
        print("ANALISIS POLA PENYEWAAN DAN OPERASIONAL PUSAT OLAHRAGA")
        print("=" * 68)

        tables = baca_database(conn)
        data_bersih, audit = bersihkan_data(tables)
        hasil = analisis_data(tables, data_bersih)
        ringkasan = buat_ringkasan(tables, hasil, audit)

        # Validasi jumlah record. Tabel bookings sendiri berisi lebih dari 1.000 baris.
        jumlah_tabel = {
            nama: len(tables[nama]) for nama in
            ["customers", "sports", "courts", "bookings",
             "booking_details", "payments", "reviews"]
        }
        total_record = sum(jumlah_tabel.values())
        print("\nJUMLAH RECORD PER TABEL")
        for nama, jumlah in jumlah_tabel.items():
            print(f"{nama:18}: {jumlah:,}")
        print(f"Total seluruh tabel: {total_record:,}")
        print(f"Record booking utama: {jumlah_tabel['bookings']:,}")

        print("\nPEMERIKSAAN KUALITAS DATA")
        print(audit.to_string(index=False))

        print("\nRINGKASAN HASIL")
        print(ringkasan.to_string(index=False))
        for nama, df in hasil.items():
            print(f"\n{nama.upper().replace('_', ' ')}")
            print(df.to_string(index=False))

        # CSV berisi data rinci hasil JOIN yang sudah dibersihkan.
        csv_data = data_bersih.copy()
        csv_data["booking_date"] = csv_data["booking_date"].dt.strftime("%Y-%m-%d")
        csv_data.to_csv(CSV_FILE, index=False, encoding="utf-8-sig")

        # Excel berisi ringkasan, audit, setiap hasil analisis, dan data rinci.
        sheet_data = {
            "Ringkasan": ringkasan,
            "Kualitas Data": audit,
            "Status Booking": hasil["status_booking"],
            "Kinerja Olahraga": hasil["kinerja_olahraga"],
            "Jam Ramai": hasil["jam_ramai"],
            "Hari Ramai": hasil["hari_ramai"],
            "Pendapatan Bulanan": hasil["pendapatan_bulanan"],
            "Metode Pembayaran": hasil["metode_pembayaran"],
            "Segmen Pelanggan": hasil["segmen_pelanggan"],
            "Tipe Lapangan": hasil["tipe_lapangan"],
            "Rating Olahraga": hasil["rating_olahraga"],
            "Pembatalan Bulanan": hasil["pembatalan_bulanan"],
            "Data Rinci": csv_data
        }
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
            for nama_sheet, dataframe in sheet_data.items():
                dataframe.to_excel(writer, sheet_name=nama_sheet, index=False)

        buat_grafik(hasil)

        print("\nKESIMPULAN SINGKAT")
        print(
            f"Cabang dengan booking selesai dan pembayaran terbesar jumlahnya adalah "
            f"{ringkasan.iloc[5, 1]}."
        )
        print(
            f"Jam paling ramai dimulai pukul {ringkasan.iloc[6, 1]}, "
            f"sedangkan pendapatan tertinggi terjadi pada {ringkasan.iloc[7, 1]}."
        )
        print(
            f"Tingkat pembatalan keseluruhan adalah "
            f"{ringkasan.iloc[3, 1]}%."
        )
        print("\nOUTPUT BERHASIL DIBUAT")
        print(f"Grafik      : {GRAFIK_DIR}")
        print(f"Excel       : {EXCEL_FILE}")
        print(f"CSV         : {CSV_FILE}")
    finally:
        conn.close()

if __name__ == "__main__":
    main()
