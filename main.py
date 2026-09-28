import io
import re
import json
import tempfile
import os
from pathlib import Path

import pdfplumber
import pikepdf
import pandas as pd
from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI(title="PDF Verifier")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

app.mount("/static", StaticFiles(directory="static"), name="static")


# ── helpers ──────────────────────────────────────────────────────────────────

def clean_number(text: str) -> float | None:
    """Ubah string angka format Indonesia ke float. Contoh: '53.895.000,00' → 53895000.0"""
    if not text:
        return None
    text = text.strip().replace(" ", "")
    # Hapus prefix mata uang
    text = re.sub(r"^(Rp\.?|IDR)", "", text, flags=re.IGNORECASE).strip()
    # Format Indonesia: titik ribuan, koma desimal
    if re.match(r"^[\d.]+,\d+$", text):
        text = text.replace(".", "").replace(",", ".")
    else:
        text = text.replace(",", "")
    try:
        return float(text)
    except ValueError:
        return None


def extract_pdf_data(pdf_bytes: bytes, password: str) -> dict:
    """
    Buka PDF terenkripsi lalu ekstrak field penting:
    - nomor_invoice / nomor_dokumen
    - vendor / supplier name
    - tanggal
    - dpp, pph, total, ppn
    Semua nilai sebagai string mentah + versi numerik.
    """
    # 1. Decrypt dengan pikepdf
    try:
        with pikepdf.open(io.BytesIO(pdf_bytes), password=password) as src:
            out_buf = io.BytesIO()
            src.save(out_buf)
            out_buf.seek(0)
            unlocked_bytes = out_buf.read()
    except pikepdf.PasswordError:
        raise HTTPException(status_code=400, detail="Password PDF salah atau file rusak")
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Gagal membuka PDF: {e}")

    # 2. Ekstrak teks
    full_text = ""
    with pdfplumber.open(io.BytesIO(unlocked_bytes)) as pdf:
        for page in pdf.pages:
            t = page.extract_text() or ""
            full_text += t + "\n"

    data = {"raw_text": full_text, "fields": {}}
    f = data["fields"]

    # ── Nomor invoice / dokumen ───────────────────────────────────────────
    patterns_inv = [
        r"(?:No\.?\s*Invoice|Invoice\s*No\.?|Nomor\s*Invoice|No\.\s*Faktur|Nomor\s*Faktur|KUITANSI\s*NO\.?|Kuitansi\s*No\.?)\s*[:\-]?\s*([A-Z0-9/_\-\.]+)",
        r"(?:INV|KWT|KWI|NS|WPLI|TMI|MGX|SI)\s*[-/]?\s*([A-Z0-9/_\-\.]+)",
    ]
    for pat in patterns_inv:
        m = re.search(pat, full_text, re.IGNORECASE)
        if m:
            f["nomor_dokumen"] = m.group(1).strip()
            break

    # ── Nomor Faktur Pajak ────────────────────────────────────────────────
    m = re.search(r"(0[48]\d{14})", full_text)
    if m:
        f["nomor_faktur_pajak"] = m.group(1)

    # ── Tanggal ───────────────────────────────────────────────────────────
    m = re.search(
        r"(\d{1,2}[\s/\-](?:Januari|Februari|Maret|April|Mei|Juni|Juli|Agustus|September|Oktober|November|Desember|January|February|March|April|May|June|July|August|September|October|November|December)[\s/\-]\d{4})",
        full_text, re.IGNORECASE
    )
    if not m:
        m = re.search(r"(\d{2}[/\-]\d{2}[/\-]\d{4})", full_text)
    if m:
        f["tanggal"] = m.group(1).strip()

    # ── Angka penting ─────────────────────────────────────────────────────
    # DPP
    m = re.search(r"DPP[^0-9]*([\d.,]+)", full_text, re.IGNORECASE)
    if m:
        f["dpp_str"] = m.group(1)
        f["dpp"] = clean_number(m.group(1))

    # PPH / PPh
    m = re.search(r"PP[Hh]\s*23?[^0-9]*([\d.,]+)", full_text)
    if not m:
        m = re.search(r"(?:Potongan|Pemotongan)\s*PP[Hh][^0-9]*([\d.,]+)", full_text, re.IGNORECASE)
    if m:
        f["pph_str"] = m.group(1)
        f["pph"] = clean_number(m.group(1))

    # Total / Jumlah
    for label in ["TOTAL RP", "TOTAL Rp", "TOTAL TAGIHAN", "JUMLAH TAGIHAN", "GRAND TOTAL", "Total Tagihan"]:
        m = re.search(re.escape(label) + r"[^0-9]*([\d.,]+)", full_text, re.IGNORECASE)
        if m:
            f["total_str"] = m.group(1)
            f["total"] = clean_number(m.group(1))
            break

    # PPN
    m = re.search(r"PPN[^0-9]*([\d.,]+)", full_text, re.IGNORECASE)
    if m:
        f["ppn_str"] = m.group(1)
        f["ppn"] = clean_number(m.group(1))

    # ── Vendor name (baris pertama / header) ─────────────────────────────
    lines = [l.strip() for l in full_text.split("\n") if l.strip()]
    for line in lines[:10]:
        if len(line) > 8 and not re.match(r"^[\d/\-]+$", line):
            f["vendor_raw"] = line
            break

    return data


def read_excel(excel_bytes: bytes) -> pd.DataFrame:
    df = pd.read_excel(io.BytesIO(excel_bytes), dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def find_excel_match(pdf_fields: dict, df: pd.DataFrame) -> dict:
    """
    Cari baris di Excel yang paling cocok dengan data PDF.
    Matching: nomor_dokumen, nomor_faktur_pajak → lalu cek DPP & PPH.
    """
    result = {
        "matched_row": None,
        "row_index": None,
        "checks": {},
        "status": "NOT_FOUND",
    }

    search_keys = []
    if "nomor_faktur_pajak" in pdf_fields:
        search_keys.append(pdf_fields["nomor_faktur_pajak"])
    if "nomor_dokumen" in pdf_fields:
        search_keys.append(pdf_fields["nomor_dokumen"])

    matched_row = None
    matched_idx = None

    for key in search_keys:
        key_clean = key.strip().upper()
        for idx, row in df.iterrows():
            row_str = " ".join(str(v) for v in row.values).upper()
            if key_clean in row_str:
                matched_row = row
                matched_idx = idx
                break
        if matched_row is not None:
            break

    if matched_row is None:
        result["status"] = "NOT_FOUND"
        return result

    result["matched_row"] = matched_row.to_dict()
    result["row_index"] = int(matched_idx) + 2  # +2 karena header + 0-index

    # Cek DPP
    checks = {}
    if "dpp" in pdf_fields and pdf_fields["dpp"]:
        excel_dpp = None
        for col in df.columns:
            if "dpp" in col.lower():
                val = clean_number(str(matched_row.get(col, "")))
                if val:
                    excel_dpp = val
                    break
        if excel_dpp is not None:
            diff = abs(pdf_fields["dpp"] - excel_dpp)
            checks["DPP"] = {
                "pdf": pdf_fields["dpp"],
                "excel": excel_dpp,
                "match": diff < 5,  # toleransi Rp5 pembulatan
                "diff": diff,
            }

    # Cek PPH
    if "pph" in pdf_fields and pdf_fields["pph"]:
        excel_pph = None
        for col in df.columns:
            if "pph" in col.lower() or "pajak" in col.lower():
                val = clean_number(str(matched_row.get(col, "")))
                if val:
                    excel_pph = val
                    break
        if excel_pph is not None:
            diff = abs(pdf_fields["pph"] - excel_pph)
            checks["PPH"] = {
                "pdf": pdf_fields["pph"],
                "excel": excel_pph,
                "match": diff < 5,
                "diff": diff,
            }

    # Cek Nomor Dokumen (string)
    if search_keys:
        for col in df.columns:
            if any(k in col.lower() for k in ["nomor", "no.", "invoice", "faktur"]):
                val = str(matched_row.get(col, "")).strip().upper()
                for key in search_keys:
                    if key.upper() in val or val in key.upper():
                        checks["Nomor Dokumen"] = {
                            "pdf": key,
                            "excel": str(matched_row.get(col, "")),
                            "match": True,
                            "diff": 0,
                        }
                        break

    result["checks"] = checks
    all_match = all(c["match"] for c in checks.values()) if checks else True
    result["status"] = "OK" if all_match else "MISMATCH"
    return result


# ── routes ───────────────────────────────────────────────────────────────────

@app.get("/", response_class=HTMLResponse)
async def root():
    html_path = Path("static/index.html")
    if html_path.exists():
        return HTMLResponse(html_path.read_text(encoding="utf-8"))
    return HTMLResponse("<h2>Static files not found</h2>")


@app.post("/verify")
async def verify(
    pdfs: list[UploadFile] = File(...),
    excel: UploadFile = File(...),
    password: str = Form(""),
):
    # Baca Excel
    excel_bytes = await excel.read()
    try:
        df = read_excel(excel_bytes)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Gagal baca Excel: {e}")

    results = []

    for pdf_file in pdfs:
        pdf_bytes = await pdf_file.read()
        entry = {
            "filename": pdf_file.filename,
            "status": "ERROR",
            "pdf_fields": {},
            "match": {},
            "error": None,
        }
        try:
            data = extract_pdf_data(pdf_bytes, password)
            entry["pdf_fields"] = data["fields"]
            entry["match"] = find_excel_match(data["fields"], df)
            entry["status"] = entry["match"]["status"]
        except HTTPException as e:
            entry["error"] = e.detail
        except Exception as e:
            entry["error"] = str(e)

        results.append(entry)

    return JSONResponse({"results": results, "excel_columns": list(df.columns)})


@app.get("/health")
async def health():
    return {"status": "ok"}
