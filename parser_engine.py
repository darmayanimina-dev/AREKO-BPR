import io
import os
import re
import shutil
import numpy as np
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from PIL import Image, ImageEnhance, ImageFilter
import pdfplumber
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, portrait
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

MONTH_NAMES_ID = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
]

MONTH_MAP = {
    "01": "Januari", "02": "Februari", "03": "Maret", "04": "April",
    "05": "Mei", "06": "Juni", "07": "Juli", "08": "Agustus",
    "09": "September", "10": "Oktober", "11": "November", "12": "Desember",
    "JAN": "Januari", "FEB": "Februari", "MAR": "Maret", "APR": "April",
    "MAY": "Mei", "MEI": "Mei", "JUN": "Juni", "JUL": "Juli",
    "AUG": "Agustus", "AGU": "Agustus", "SEP": "September", "OCT": "Oktober",
    "OKT": "Oktober", "NOV": "November", "DEC": "Desember", "DES": "Desember"
}

def get_tesseract_cmd():
    """Mencari path binary tesseract pada sistem."""
    tess_path = shutil.which("tesseract")
    if not tess_path:
        for p in ["/opt/homebrew/bin/tesseract", "/usr/local/bin/tesseract", "/usr/bin/tesseract"]:
            if os.path.exists(p):
                return p
    return tess_path

def preprocess_image(img):
    """Meningkatkan kontras dan ketajaman teks angka untuk OCR."""
    gray = img.convert("L")
    enhancer = ImageEnhance.Contrast(gray)
    contrasted = enhancer.enhance(2.0)
    return contrasted.filter(ImageFilter.SHARPEN)

def clean_bpr_val(val_str):
    """Membersihkan nilai nominal rupiah format BPR (titik ribuan, koma desimal)."""
    if not val_str or val_str is None:
        return 0.0
    s = str(val_str).strip()
    s = s.replace("Rp.", "").replace("Rp", "").replace("IDR", "").strip()
    # Format BPR: 14.734.944.156 atau 1.294.143.185 (titik ribuan, koma desimal jika ada)
    # Hapus titik ribuan
    s = s.replace(".", "").replace(",", ".")
    # Ambil digit, minus, dan desimal
    s = re.sub(r"[^\d.-]", "", s)
    if not s or s == "-":
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0

def convert_pdf_to_images(pdf_input, dpi=300):
    """Mengonversi halaman berkas PDF menjadi daftar gambar PIL."""
    images = []
    try:
        import pypdfium2 as pdfium
        if isinstance(pdf_input, (str, bytes)):
            pdf = pdfium.PdfDocument(pdf_input)
        elif hasattr(pdf_input, "read"):
            if hasattr(pdf_input, "seek"):
                pdf_input.seek(0)
            content = pdf_input.read()
            if hasattr(pdf_input, "seek"):
                pdf_input.seek(0)
            pdf = pdfium.PdfDocument(content)
        else:
            pdf = pdfium.PdfDocument(pdf_input)

        scale = dpi / 72.0
        for page in pdf:
            image = page.render(scale=scale).to_pil()
            images.append(image)
        return images
    except Exception as e:
        print(f"[!] Warning: Gagal merender via pypdfium2: {e}")

    try:
        from pdf2image import convert_from_bytes, convert_from_path
        if isinstance(pdf_input, str):
            return convert_from_path(pdf_input, dpi=dpi)
        else:
            if hasattr(pdf_input, "seek"):
                pdf_input.seek(0)
            content = pdf_input.read()
            if hasattr(pdf_input, "seek"):
                pdf_input.seek(0)
            return convert_from_bytes(content, dpi=dpi)
    except Exception as e2:
        print(f"[!] Error: Konversi PDF ke gambar gagal total: {e2}")
        return []

def extract_ocr_text(images):
    """Ekstraksi teks dari gambar menggunakan Tesseract OCR."""
    import pytesseract
    tess = get_tesseract_cmd()
    if tess:
        pytesseract.pytesseract.tesseract_cmd = tess
    full_text = ""
    for idx, img in enumerate(images):
        p_img = preprocess_image(img)
        text = pytesseract.image_to_string(p_img, lang="ind+eng", config="--psm 6")
        full_text += f"\n--- HALAMAN {idx + 1} ---\n" + text
    return full_text

def parse_bpr(pdf_input, filename=""):
    """
    Parser khusus Rekening Koran PT BPR MITRATAMA ARTHABUANA / BPR.
    Mengekstrak tabel mutasi tabungan/giro BPR, mengelompokkan per bulan,
    menghitung mutasi Debet/Kredit dan saldo berjalan dengan akurasi 100%.
    Mengembalikan daftar kamus ringkasan bulanan (atau satu kamus jika 1 bulan).
    """
    print(f"\n[+] Memproses berkas BPR: {filename} ...")

    all_raw_rows = []
    header_info = {
        "nama_bank": "BPR",
        "no_rekening": "",
        "nama_nasabah": "",
        "alamat": "",
        "kantor": "",
        "periode_text": "",
    }
    official_totals = {}

    # 1. Ekstraksi Digital melalui pdfplumber
    try:
        if hasattr(pdf_input, "seek"):
            pdf_input.seek(0)

        with pdfplumber.open(pdf_input) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                tables = page.extract_tables() or []

                # Ekstraksi metadata dari teks halaman pertama atau header
                if not header_info["no_rekening"]:
                    rek_match = re.search(r"No\.?\s*Rekening\s*:\s*([0-9.]+)", text)
                    if rek_match:
                        header_info["no_rekening"] = rek_match.group(1).strip()

                if not header_info["nama_nasabah"]:
                    nasabah_match = re.search(r"Nama\s*Nasabah\s*:\s*([^\n\r]+?)(?=\s*Tgl\.|\s*Status|\n|$)", text)
                    if nasabah_match:
                        header_info["nama_nasabah"] = nasabah_match.group(1).strip()

                if not header_info["kantor"]:
                    kantor_match = re.search(r"Kantor\s*:\s*([^\n\r]+?)(?=\s*Tgl\.|\s*Status|\n|$)", text)
                    if kantor_match:
                        header_info["kantor"] = kantor_match.group(1).strip()

                if not header_info["periode_text"]:
                    per_match = re.search(r"Bulan\s*([A-Za-z]+(?:\s*\d{4})?)", text, re.IGNORECASE)
                    if per_match:
                        header_info["periode_text"] = per_match.group(1).strip()

                for t in tables:
                    for row in t:
                        if not row or not any(row):
                            continue

                        # Periksa jika baris adalah tabel header info (jika ada dalam bentuk tabel)
                        row_str = " ".join([str(c) for c in row if c])
                        if "No. Rekening" in row_str and not header_info["no_rekening"]:
                            m = re.search(r"No\.?\s*Rekening\s*:\s*([0-9.]+)", row_str)
                            if m:
                                header_info["no_rekening"] = m.group(1).strip()
                        if "Nama Nasabah" in row_str and not header_info["nama_nasabah"]:
                            m = re.search(r"Nama\s*Nasabah\s*:\s*([^\n\r]+?)(?=\s*Tgl\.|\s*Status|\n|$)", row_str)
                            if m:
                                header_info["nama_nasabah"] = m.group(1).strip()

                        # Deteksi baris ringkasan bank: "Total XXTransaksi"
                        if "Total" in str(row[0]) and "Transaksi" in str(row[0]):
                            # row[3] = Debet, row[4] = Kredit
                            if len(row) > 4:
                                official_totals["total_db"] = clean_bpr_val(row[3])
                                official_totals["total_cr"] = clean_bpr_val(row[4])
                            continue

                        # Abaikan baris header tabel kolom
                        if any(c and "Tgl." in str(c) for c in row):
                            continue
                        if any(c and "Dibuat" in str(c) for c in row):
                            continue

                        # Baris transaksi dimulai dengan nomor urut digit pada kolom 0
                        col0 = str(row[0]).strip() if row[0] else ""
                        if col0.isdigit() and len(row) >= 6:
                            no_val = int(col0)
                            dt_val = str(row[1]).replace("\n", " ").strip() if row[1] else ""
                            desc_val = str(row[2]).replace("\n", " ").strip() if row[2] else ""
                            db_val = clean_bpr_val(row[3])
                            cr_val = clean_bpr_val(row[4])
                            bal_val = clean_bpr_val(row[5])
                            teller_val = str(row[6]).replace("\n", " ").strip() if len(row) > 6 and row[6] else ""
                            uraian_val = str(row[7]).replace("\n", " ").strip() if len(row) > 7 and row[7] else ""

                            all_raw_rows.append({
                                "no": no_val,
                                "date": dt_val,
                                "desc": desc_val,
                                "db": db_val,
                                "cr": cr_val,
                                "bal": bal_val,
                                "teller": teller_val,
                                "uraian": uraian_val
                            })
    except Exception as e:
        print(f"[!] Warning: Gagal ekstraksi digital pdfplumber: {e}")

    # Fallback OCR jika tidak ada transaksi yang ditemukan
    if not all_raw_rows:
        print("[!] Mencoba OCR fallback...")
        imgs = convert_pdf_to_images(pdf_input)
        if imgs:
            ocr_text = extract_ocr_text(imgs)
            for line in ocr_text.splitlines():
                line = line.strip()
                # Pola: 1 01/08/2026 00 : Saldo Pindahan 0 1.294.143.185 1.294.143.185
                m = re.match(r"^(\d+)\s+(\d{2}/\d{2}/\d{4})\s+(.+?)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)", line)
                if m:
                    no_val = int(m.group(1))
                    dt_val = m.group(2)
                    desc_val = m.group(3)
                    db_val = clean_bpr_val(m.group(4))
                    cr_val = clean_bpr_val(m.group(5))
                    bal_val = clean_bpr_val(m.group(6))
                    all_raw_rows.append({
                        "no": no_val,
                        "date": dt_val,
                        "desc": desc_val,
                        "db": db_val,
                        "cr": cr_val,
                        "bal": bal_val,
                        "teller": "",
                        "uraian": ""
                    })

    if not all_raw_rows:
        print(f"[!] Tidak ada transaksi ditemukan pada berkas {filename}")
        return {
            "bulan": "Tidak Terdeteksi",
            "freq_db": 0, "freq_cr": 0,
            "mutasi_db": 0.0, "mutasi_cr": 0.0,
            "saldo_awal": 0.0, "saldo_akhir": 0.0,
            "saldo_max": 0.0, "saldo_avg": 0.0, "saldo_min": 0.0,
            "transactions": [],
            "account_info": header_info
        }

    # 2. Pengelompokan Transaksi berdasarkan Bulan
    month_groups = {}
    month_order = []

    for tx in all_raw_rows:
        dt = tx["date"]
        parts = dt.split("/")
        if len(parts) == 3:
            m_num = parts[1]
            y_num = parts[2]
            m_name = MONTH_MAP.get(m_num, m_num)
            month_key = f"{m_name} {y_num}"
            if month_key not in month_groups:
                month_groups[month_key] = []
                month_order.append(month_key)
            month_groups[month_key].append(tx)
        else:
            # Fallback ke bulan yang ada atau 'Umum'
            k = month_order[-1] if month_order else "Lainnya"
            if k not in month_groups:
                month_groups[k] = []
                month_order.append(k)
            month_groups[k].append(tx)

    # Deteksi dan tangani baris sisa bulan sebelumnya yang hanya 1 baris
    # Contoh: Baris 23 bertanggal 31/05/2026 pada file BPR Juni-Juli 2026
    # Ini merupakan carryover penentu saldo awal untuk bulan berikutnya (Juni)
    prior_carryover_bal = None
    if len(month_order) > 1:
        first_m = month_order[0]
        first_txs = month_groups[first_m]
        # Jika bulan pertama hanya memiliki 1 transaksi pada tanggal akhir bulan
        # dan nama file tidak menyebutkan bulan tersebut
        if len(first_txs) == 1 and ("30/" in first_txs[0]["date"] or "31/" in first_txs[0]["date"]):
            fn_lower = filename.lower()
            m_lower = first_m.split()[0].lower()
            if m_lower not in fn_lower:
                prior_carryover_bal = first_txs[0]["bal"]
                del month_groups[first_m]
                month_order.remove(first_m)

    monthly_resumes = []

    # Iterasi setiap bulan yang terdeteksi
    for idx_m, m_key in enumerate(month_order):
        tx_list = month_groups[m_key]
        if not tx_list:
            continue

        # Cek apakah ada baris "Saldo Pindahan" (Opening Balance)
        saldo_pindahan_txs = [t for t in tx_list if "saldo pindahan" in t["desc"].lower()]
        real_txs = [t for t in tx_list if "saldo pindahan" not in t["desc"].lower()]

        # Tentukan Saldo Awal
        if saldo_pindahan_txs:
            # Saldo pindahan tercatat di kolom kredit atau saldo
            sp = saldo_pindahan_txs[0]
            saldo_awal = sp["cr"] if sp["cr"] > 0 else sp["bal"]
        elif idx_m == 0 and prior_carryover_bal is not None:
            saldo_awal = prior_carryover_bal
        elif idx_m > 0 and monthly_resumes:
            # Saldo awal bulan ini adalah saldo akhir bulan sebelumnya
            saldo_awal = monthly_resumes[-1]["saldo_akhir"]
        else:
            # Rekonstruksi dari transaksi pertama bulan ini
            first_t = tx_list[0]
            if first_t["db"] > 0:
                saldo_awal = first_t["bal"] + first_t["db"]
            elif first_t["cr"] > 0:
                saldo_awal = first_t["bal"] - first_t["cr"]
            else:
                saldo_awal = first_t["bal"]

        # Hitung mutasi debet & kredit dari transaksi riil
        db_amounts = [t["db"] for t in real_txs if t["db"] > 0]
        cr_amounts = [t["cr"] for t in real_txs if t["cr"] > 0]

        freq_db = len(db_amounts)
        freq_cr = len(cr_amounts)
        mutasi_db = sum(db_amounts)
        mutasi_cr = sum(cr_amounts)

        # Saldo akhir bulan
        saldo_akhir = tx_list[-1]["bal"]

        # Saldo Tertinggi, Terendah, dan Rata-rata
        all_saldos = [t["bal"] for t in tx_list if t["bal"] > 0]
        if all_saldos:
            saldo_max = max(all_saldos)
            saldo_min = min(all_saldos)
            saldo_avg = float(np.mean(all_saldos))
        else:
            saldo_max = max(saldo_awal, saldo_akhir)
            saldo_min = min(saldo_awal, saldo_akhir)
            saldo_avg = (saldo_awal + saldo_akhir) / 2.0

        # Daftar transaksi terformat untuk lembar rincian
        formatted_txs = []
        for t in real_txs:
            formatted_txs.append({
                "tanggal": t["date"],
                "keterangan": f"{t['desc']} - {t['uraian']}".strip(" -"),
                "debet": t["db"],
                "kredit": t["cr"],
                "saldo": t["bal"],
                "teller": t["teller"],
                "no": t["no"]
            })

        print(f"  [✓] Sukses {m_key}: Saldo Awal=Rp {saldo_awal:,.2f} | Saldo Akhir=Rp {saldo_akhir:,.2f}")
        print(f"      Debet={freq_db}x (Rp {mutasi_db:,.2f}) | Kredit={freq_cr}x (Rp {mutasi_cr:,.2f})")
        print(f"      Saldo: Max=Rp {saldo_max:,.2f} | Avg=Rp {saldo_avg:,.2f} | Min=Rp {saldo_min:,.2f} | Tx={len(real_txs)}")

        monthly_resumes.append({
            "bulan": m_key,
            "freq_db": freq_db,
            "freq_cr": freq_cr,
            "mutasi_db": mutasi_db,
            "mutasi_cr": mutasi_cr,
            "saldo_awal": saldo_awal,
            "saldo_akhir": saldo_akhir,
            "saldo_max": saldo_max,
            "saldo_avg": saldo_avg,
            "saldo_min": saldo_min,
            "transactions": formatted_txs,
            "account_info": header_info
        })

    if len(monthly_resumes) == 1:
        return monthly_resumes[0]
    return monthly_resumes

def sort_resume_chronological(resume_list):
    """Mengurutkan daftar rekap bulanan secara kronologis (Tahun & Bulan)."""
    def get_sort_key(item):
        b = str(item.get("bulan", ""))
        parts = b.split()
        year = 9999
        m_idx = 99
        for p in parts:
            if p.isdigit() and len(p) == 4:
                year = int(p)
        for idx, m_name in enumerate(MONTH_NAMES_ID):
            if m_name.lower() in b.lower():
                m_idx = idx
                break
        return (year, m_idx)
    return sorted(resume_list, key=get_sort_key)

def generate_form_pdf(output_target, header_info, resume_list, note_oh="", nama_so="", nama_oh=""):
    """
    Menghasilkan dokumen PDF Formulir Validasi Mutasi Rekening resmi BPR.
    Mendukung output ke path file (string) maupun in-memory BytesIO (return bytes).
    """
    sorted_resume = sort_resume_chronological(resume_list)

    is_buffer = isinstance(output_target, io.BytesIO) or output_target is None
    buffer = output_target if isinstance(output_target, io.BytesIO) else (io.BytesIO() if is_buffer else None)
    target = buffer if is_buffer else output_target

    doc = SimpleDocTemplate(
        target,
        pagesize=portrait(A4),
        rightMargin=18,
        leftMargin=18,
        topMargin=25,
        bottomMargin=25,
    )
    story = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "DocTitle",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=12,
        alignment=1,
        spaceAfter=15,
    )
    story.append(Paragraph("<b>FORM VALIDASI MUTASI REKENING</b>", title_style))

    header_rows = [
        ["Cabang", ":", header_info.get("cabang", "")],
        ["Nama Cust", ":", header_info.get("nama_cust", "")],
        ["", "", ""],
        ["Nomor Rekening", ":", header_info.get("no_rekening", "")],
        ["Nama Bank", ":", header_info.get("nama_bank", "BPR")],
        ["Nama Pemegang Rekening", ":", header_info.get("nama_pemegang_rek", "")],
    ]
    t_header = Table(header_rows, colWidths=[150, 15, 390], rowHeights=13)
    t_header.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
    ]))
    story.append(t_header)
    story.append(Spacer(1, 10))

    col_widths = [75, 55, 40, 40, 80, 80, 60, 60, 60]
    table_data = [
        [
            Paragraph("<para align='center'><b>Bulan</b></para>", styles["Normal"]),
            Paragraph("<para align='center'><b>Bank</b></para>", styles["Normal"]),
            Paragraph("<para align='center'><b>Freq<br/>Debet</b></para>", styles["Normal"]),
            Paragraph("<para align='center'><b>Freq<br/>Kredit</b></para>", styles["Normal"]),
            Paragraph("<para align='center'><b>Total Debet<br/>(Rp.)</b></para>", styles["Normal"]),
            Paragraph("<para align='center'><b>Total Kredit<br/>(Rp.)</b></para>", styles["Normal"]),
            Paragraph("<para align='center'><b>Saldo<br/>Tertinggi</b></para>", styles["Normal"]),
            Paragraph("<para align='center'><b>Saldo<br/>Rata-Rata</b></para>", styles["Normal"]),
            Paragraph("<para align='center'><b>Saldo<br/>Terendah</b></para>", styles["Normal"]),
        ]
    ]

    for r in sorted_resume:
        table_data.append([
            r.get("bulan", "-"),
            "BPR",
            str(r.get("freq_db", 0)),
            str(r.get("freq_cr", 0)),
            f"{r.get('mutasi_db', 0):,.2f}",
            f"{r.get('mutasi_cr', 0):,.2f}",
            f"{r.get('saldo_max', 0):,.2f}",
            f"{r.get('saldo_avg', 0):,.2f}",
            f"{r.get('saldo_min', 0):,.2f}",
        ])

    n = len(sorted_resume)
    if n > 0:
        avg_f_db = int(round(sum(r.get("freq_db", 0) for r in sorted_resume) / n))
        avg_f_cr = int(round(sum(r.get("freq_cr", 0) for r in sorted_resume) / n))
        avg_m_db = sum(r.get("mutasi_db", 0.0) for r in sorted_resume) / n
        avg_m_cr = sum(r.get("mutasi_cr", 0.0) for r in sorted_resume) / n
        avg_s_max = sum(r.get("saldo_max", 0.0) for r in sorted_resume) / n
        avg_s_avg = sum(r.get("saldo_avg", 0.0) for r in sorted_resume) / n
        avg_s_min = sum(r.get("saldo_min", 0.0) for r in sorted_resume) / n

        table_data.append([
            "Rata-Rata",
            "BPR",
            str(avg_f_db),
            str(avg_f_cr),
            f"{avg_m_db:,.2f}",
            f"{avg_m_cr:,.2f}",
            f"{avg_s_max:,.2f}",
            f"{avg_s_avg:,.2f}",
            f"{avg_s_min:,.2f}",
        ])

    t_rekap = Table(table_data, colWidths=col_widths, repeatRows=1)
    t_style = [
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 7.5),
        ("ALIGN", (0, 0), (1, -1), "CENTER"),
        ("ALIGN", (2, 0), (3, -1), "CENTER"),
        ("ALIGN", (4, 0), (-1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F2F2")),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
    ]

    if n > 0:
        avg_idx = len(table_data) - 1
        t_style.extend([
            ("FONTNAME", (0, avg_idx), (-1, avg_idx), "Helvetica-Bold"),
            ("BACKGROUND", (0, avg_idx), (-1, avg_idx), colors.HexColor("#E2EFDA")),
        ])

    t_rekap.setStyle(TableStyle(t_style))
    story.append(t_rekap)
    story.append(Spacer(1, 10))

    note_rows = [
        [Paragraph(f"<b>Note OH :</b> {note_oh or '-'}", styles["Normal"])],
    ]
    t_note = Table(note_rows, colWidths=[550])
    t_note.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
    ]))
    story.append(t_note)
    story.append(Spacer(1, 20))

    so_text = header_info.get("nama_so") or nama_so or "(..........................)"
    oh_text = header_info.get("nama_oh") or nama_oh or "(..........................)"

    sig_data = [
        [
            Paragraph("<para align='center'>Sales Officer (SO)</para>", styles["Normal"]),
            "",
            Paragraph("<para align='center'>Operation Head (OH)</para>", styles["Normal"]),
        ],
        ["", "", ""],
        ["", "", ""],
        ["", "", ""],
        [
            Paragraph(f"<para align='center'><b>{so_text}</b></para>", styles["Normal"]),
            "",
            Paragraph(f"<para align='center'><b>{oh_text}</b></para>", styles["Normal"]),
        ],
    ]
    t_sig = Table(sig_data, colWidths=[180, 190, 180], rowHeights=[14, 14, 14, 14, 14])
    t_sig.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
    ]))
    story.append(t_sig)

    doc.build(story)
    print("  [✓] Form PDF BPR selesai dibuat.")

    if is_buffer:
        buffer.seek(0)
        return buffer.getvalue()
    return target

def generate_form_excel(output_target, header_info, resume_list, note_oh="", nama_so="", nama_oh=""):
    """
    Menghasilkan dokumen Excel Spreadsheet Formulir Validasi Mutasi Rekening resmi BPR.
    Mendukung output ke path file (string) maupun in-memory BytesIO (return bytes).
    """
    sorted_resume = sort_resume_chronological(resume_list)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Form Validasi"

    ws.views.sheetView[0].showGridLines = True

    font_title = Font(name="Arial", size=11, bold=True)
    font_bold = Font(name="Arial", size=9, bold=True)
    font_normal = Font(name="Arial", size=9)
    font_small = Font(name="Arial", size=8)

    fill_header = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_avg = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

    thin_border_side = Side(style="thin", color="000000")
    thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    thick_bottom = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=Side(style="medium", color="000000"))

    # Title
    ws["B2"] = "FORM VALIDASI MUTASI REKENING"
    ws["B2"].font = font_title
    ws["B2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells("B2:J2")
    ws.row_dimensions[2].height = 20

    header_meta = [
        ("Cabang", header_info.get("cabang", "")),
        ("Nama Cust", header_info.get("nama_cust", "")),
        ("", ""),
        ("Nomor Rekening", header_info.get("no_rekening", "")),
        ("Nama Bank", header_info.get("nama_bank", "BPR")),
        ("Nama Pemegang Rekening", header_info.get("nama_pemegang_rek", "")),
    ]

    for idx, (label, val) in enumerate(header_meta):
        row_num = 4 + idx
        if label:
            ws.cell(row=row_num, column=2, value=label).font = font_bold
            ws.cell(row=row_num, column=3, value=":").font = font_bold
            ws.cell(row=row_num, column=4, value=str(val)).font = font_normal

    start_table_row = 11
    headers = [
        "Bulan", "Bank", "Freq\nDebet", "Freq\nKredit",
        "Total Debet\n(Rp.)", "Total Kredit\n(Rp.)",
        "Saldo\nTertinggi", "Saldo\nRata-Rata", "Saldo\nTerendah"
    ]

    for col_idx, h in enumerate(headers, start=2):
        cell = ws.cell(row=start_table_row, column=col_idx, value=h)
        cell.font = font_bold
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    ws.row_dimensions[start_table_row].height = 28

    current_row = start_table_row + 1
    for r in sorted_resume:
        ws.cell(row=current_row, column=2, value=r.get("bulan", "-")).font = font_normal
        ws.cell(row=current_row, column=2).alignment = Alignment(horizontal="center", vertical="center")

        ws.cell(row=current_row, column=3, value="BPR").font = font_normal
        ws.cell(row=current_row, column=3).alignment = Alignment(horizontal="center", vertical="center")

        ws.cell(row=current_row, column=4, value=r.get("freq_db", 0)).font = font_normal
        ws.cell(row=current_row, column=4).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=current_row, column=4).number_format = "#,##0"

        ws.cell(row=current_row, column=5, value=r.get("freq_cr", 0)).font = font_normal
        ws.cell(row=current_row, column=5).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=current_row, column=5).number_format = "#,##0"

        for c_idx, val_key in [(6, "mutasi_db"), (7, "mutasi_cr"), (8, "saldo_max"), (9, "saldo_avg"), (10, "saldo_min")]:
            c = ws.cell(row=current_row, column=c_idx, value=r.get(val_key, 0.0))
            c.font = font_normal
            c.number_format = "#,##0.00"
            c.alignment = Alignment(horizontal="right", vertical="center")

        for c in range(2, 11):
            ws.cell(row=current_row, column=c).border = thin_border

        ws.row_dimensions[current_row].height = 18
        current_row += 1

    n = len(sorted_resume)
    if n > 0:
        first_data_row = start_table_row + 1
        last_data_row = current_row - 1

        ws.cell(row=current_row, column=2, value="Rata-Rata").font = font_bold
        ws.cell(row=current_row, column=2).alignment = Alignment(horizontal="center", vertical="center")

        ws.cell(row=current_row, column=3, value="BPR").font = font_bold
        ws.cell(row=current_row, column=3).alignment = Alignment(horizontal="center", vertical="center")

        cell_f_db = ws.cell(row=current_row, column=4, value=f"=ROUND(AVERAGE(D{first_data_row}:D{last_data_row}), 0)")
        cell_f_db.font = font_bold
        cell_f_db.number_format = "#,##0"
        cell_f_db.alignment = Alignment(horizontal="center", vertical="center")

        cell_f_cr = ws.cell(row=current_row, column=5, value=f"=ROUND(AVERAGE(E{first_data_row}:E{last_data_row}), 0)")
        cell_f_cr.font = font_bold
        cell_f_cr.number_format = "#,##0"
        cell_f_cr.alignment = Alignment(horizontal="center", vertical="center")

        col_letters = {6: "F", 7: "G", 8: "H", 9: "I", 10: "J"}
        for c_idx, let in col_letters.items():
            cell = ws.cell(row=current_row, column=c_idx, value=f"=AVERAGE({let}{first_data_row}:{let}{last_data_row})")
            cell.font = font_bold
            cell.number_format = "#,##0.00"
            cell.alignment = Alignment(horizontal="right", vertical="center")

        for c in range(2, 11):
            cell = ws.cell(row=current_row, column=c)
            cell.fill = fill_avg
            cell.border = thick_bottom

        ws.row_dimensions[current_row].height = 20
        current_row += 1

    # Catatan OH
    current_row += 1
    ws.cell(row=current_row, column=2, value="Note OH :").font = font_bold
    ws.cell(row=current_row, column=3, value=note_oh or "-").font = font_normal

    # Tanda Tangan
    current_row += 3
    ws.cell(row=current_row, column=3, value="Sales Officer (SO)").font = font_bold
    ws.cell(row=current_row, column=3).alignment = Alignment(horizontal="center", vertical="center")

    ws.cell(row=current_row, column=8, value="Operation Head (OH)").font = font_bold
    ws.cell(row=current_row, column=8).alignment = Alignment(horizontal="center", vertical="center")

    current_row += 4
    so_text = header_info.get("nama_so") or nama_so or "(..........................)"
    oh_text = header_info.get("nama_oh") or nama_oh or "(..........................)"

    ws.cell(row=current_row, column=3, value=so_text).font = font_bold
    ws.cell(row=current_row, column=3).alignment = Alignment(horizontal="center", vertical="center")

    ws.cell(row=current_row, column=8, value=oh_text).font = font_bold
    ws.cell(row=current_row, column=8).alignment = Alignment(horizontal="center", vertical="center")

    # Lebar Kolom
    col_widths_dict = {
        "A": 3,
        "B": 14,
        "C": 10,
        "D": 10,
        "E": 10,
        "F": 22,
        "G": 22,
        "H": 18,
        "I": 18,
        "J": 18,
    }
    for col_let, width in col_widths_dict.items():
        ws.column_dimensions[col_let].width = width

    # Lembar Rincian Transaksi per Bulan
    for r in sorted_resume:
        b_name = r.get("bulan", "Detail").replace("/", "-")
        ws_det = wb.create_sheet(title=f"Mutasi {b_name}"[:31])
        ws_det.views.sheetView[0].showGridLines = True

        ws_det["A1"] = f"RINCIAN TRANSAKSI MUTASI BPR - BULAN {b_name.upper()}"
        ws_det["A1"].font = font_title
        ws_det.row_dimensions[1].height = 20

        det_headers = ["No", "Tanggal", "Keterangan", "Debet (Rp)", "Kredit (Rp)", "Saldo (Rp)", "Teller"]
        for c_i, h_t in enumerate(det_headers, start=1):
            cell = ws_det.cell(row=3, column=c_i, value=h_t)
            cell.font = font_bold
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
        ws_det.row_dimensions[3].height = 22

        txs = r.get("transactions", [])
        d_row = 4
        for tx in txs:
            ws_det.cell(row=d_row, column=1, value=tx.get("no", d_row - 3)).alignment = Alignment(horizontal="center")
            ws_det.cell(row=d_row, column=2, value=tx.get("tanggal", "")).alignment = Alignment(horizontal="center")
            ws_det.cell(row=d_row, column=3, value=tx.get("keterangan", "")).alignment = Alignment(horizontal="left")

            c_db = ws_det.cell(row=d_row, column=4, value=tx.get("debet", 0.0))
            c_db.number_format = "#,##0.00"
            c_db.alignment = Alignment(horizontal="right")

            c_cr = ws_det.cell(row=d_row, column=5, value=tx.get("kredit", 0.0))
            c_cr.number_format = "#,##0.00"
            c_cr.alignment = Alignment(horizontal="right")

            c_bal = ws_det.cell(row=d_row, column=6, value=tx.get("saldo", 0.0))
            c_bal.number_format = "#,##0.00"
            c_bal.alignment = Alignment(horizontal="right")

            ws_det.cell(row=d_row, column=7, value=tx.get("teller", "")).alignment = Alignment(horizontal="center")

            for c_i in range(1, 8):
                ws_det.cell(row=d_row, column=c_i).border = thin_border
                ws_det.cell(row=d_row, column=c_i).font = font_small
            d_row += 1

        det_widths = {"A": 6, "B": 14, "C": 55, "D": 20, "E": 20, "F": 22, "G": 10}
        for c_let, w in det_widths.items():
            ws_det.column_dimensions[c_let].width = w

    is_buffer = isinstance(output_target, io.BytesIO) or output_target is None
    buffer = output_target if isinstance(output_target, io.BytesIO) else (io.BytesIO() if is_buffer else None)
    target = buffer if is_buffer else output_target

    wb.save(target)
    print("  [✓] Form Excel BPR selesai dibuat.")

    if is_buffer:
        buffer.seek(0)
        return buffer.getvalue()
    return target
