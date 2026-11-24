from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
import os
import socket
import sqlite3
from pydantic import BaseModel
from typing import Optional
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from backend.database import init_db, get_db_connection

app = FastAPI(title="AVKG FACTORY MANAGER")

@app.on_event("startup")
def startup_event():
    init_db()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(BASE_DIR)
FRONTEND_DIR = os.path.join(ROOT_DIR, "frontend")
EXPORTS_DIR = os.path.join(ROOT_DIR, "exports")
DB_PATH = os.path.join(ROOT_DIR, "database", "avkg_factory.db")
os.makedirs(EXPORTS_DIR, exist_ok=True)

app.mount("/static", StaticFiles(directory=FRONTEND_DIR), name="static")

@app.get("/")
def read_root():
    return FileResponse(os.path.join(FRONTEND_DIR, "index.html"))

@app.get("/api/health")
def health_check():
    hostname = socket.gethostname()
    try:
        local_ip = socket.gethostbyname(hostname)
    except:
        local_ip = "127.0.0.1"
        
    return {
        "status": "success",
        "app": "AVKG FACTORY MANAGER",
        "brand": "AVKG WORLD",
        "message": "AVKG Factory Manager Backend 100% Active hai!",
        "server_ip": f"http://{local_ip}:8000"
    }

class PartyCreate(BaseModel):
    name: str

class PartyUpdate(BaseModel):
    name: str

class StockEntryCreate(BaseModel):
    party_id: int
    entry_date: str
    lot_number: str
    size: str
    colour: str
    quantity: int
    remarks: Optional[str] = ""

class StockEntryUpdate(BaseModel):
    entry_date: str
    lot_number: str
    size: str
    colour: str
    quantity: int
    remarks: Optional[str] = ""

@app.post("/api/parties")
def create_party(party: PartyCreate):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("INSERT INTO parties (name) VALUES (?)", (party.name.strip(),))
        conn.commit()
        party_id = cursor.lastrowid
    except sqlite3.IntegrityError:
        conn.close()
        raise HTTPException(status_code=400, detail="Yeh Party name pehle se registered hai!")
    finally:
        conn.close()
    return {"status": "success", "message": "Party successfully register ho gayi!", "party_id": party_id, "name": party.name.strip()}

@app.put("/api/parties/{party_id}")
def update_party(party_id: int, party: PartyUpdate):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("UPDATE parties SET name = ? WHERE id = ?", (party.name.strip(), party_id))
        if cursor.rowcount == 0:
            raise HTTPException(status_code=404, detail="Party nahi mili!")
        conn.commit()
    except sqlite3.IntegrityError:
        conn.close()
        raise HTTPException(status_code=400, detail="Yeh Party name pehle se kisi aur party ka hai!")
    finally:
        conn.close()
    return {"status": "success", "message": "Party name update ho gaya!"}

@app.get("/api/parties")
def get_parties():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT p.id, p.name, p.created_at,
               COUNT(s.id) as total_entries,
               COALESCE(SUM(s.quantity), 0) as total_pieces,
               SUM(CASE WHEN s.is_exported = 0 THEN 1 ELSE 0 END) as unexported_count
        FROM parties p
        LEFT JOIN stock_entries s ON p.id = s.party_id
        GROUP BY p.id
        ORDER BY p.name ASC
    """)
    rows = cursor.fetchall()
    conn.close()
    
    parties_list = []
    for row in rows:
        unexported = row["unexported_count"] or 0
        parties_list.append({
            "id": row["id"],
            "name": row["name"],
            "created_at": row["created_at"],
            "total_entries": row["total_entries"],
            "total_pieces": row["total_pieces"],
            "unexported_count": unexported,
            "batch_progress": f"{unexported}/10"
        })
    return parties_list

@app.get("/api/parties/{party_id}/dashboard")
def get_party_dashboard(party_id: int):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM parties WHERE id = ?", (party_id,))
    party = cursor.fetchone()
    if not party:
        conn.close()
        raise HTTPException(status_code=404, detail="Party nahi mili!")
        
    cursor.execute("SELECT * FROM stock_entries WHERE party_id = ? ORDER BY entry_date DESC, id DESC", (party_id,))
    entries = [dict(row) for row in cursor.fetchall()]
    conn.close()
    
    total_entries = len(entries)
    total_pieces = sum(e["quantity"] for e in entries)
    chhote_pieces = sum(e["quantity"] for e in entries if e["size"] == "Chhote")
    bade_pieces = sum(e["quantity"] for e in entries if e["size"] == "Bade")
    unexported_count = sum(1 for e in entries if e["is_exported"] == 0)
    
    colour_totals = {}
    for e in entries:
        c = e["colour"]
        colour_totals[c] = colour_totals.get(c, 0) + e["quantity"]
        
    return {
        "party": dict(party),
        "stats": {
            "total_entries": total_entries,
            "total_pieces": total_pieces,
            "chhote_pieces": chhote_pieces,
            "bade_pieces": bade_pieces,
            "unexported_count": unexported_count,
            "batch_progress": f"{unexported_count}/10",
            "colour_totals": colour_totals
        },
        "entries": entries
    }

def generate_excel_batch_file(party_name: str, batch_num: int, entries: list):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Batch {batch_num:03d}" if batch_num != 999 else "Manual Export"
    ws.views.sheetView[0].showGridLines = True

    font_title = Font(name="Segoe UI", size=16, bold=True, color="FFFFFF")
    font_header = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    font_bold = Font(name="Segoe UI", size=11, bold=True)
    font_normal = Font(name="Segoe UI", size=11)
    
    fill_title = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    fill_header = PatternFill(start_color="0284C7", end_color="0284C7", fill_type="solid")
    fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    
    border_thin = Border(
        left=Side(style='thin', color='CBD5E1'), right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'), bottom=Side(style='thin', color='CBD5E1')
    )
    border_total = Border(top=Side(style='thin', color='000000'), bottom=Side(style='double', color='000000'))

    ws.merge_cells("A1:G1")
    ws["A1"] = "AVKG FACTORY MANAGER"
    ws["A1"].font = font_title
    ws["A1"].fill = fill_title
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 35

    ws["A3"] = "Party Name:"
    ws["B3"] = party_name
    ws["B3"].font = font_bold
    ws["A4"] = "Batch Number:"
    ws["B4"] = f"BATCH {batch_num:03d}" if batch_num != 999 else "MANUAL EXPORT"
    ws["B4"].font = font_bold
    ws["A5"] = "Export Date:"
    ws["B5"] = datetime.now().strftime("%d-%m-%Y")

    headers = ["S.No.", "Date", "Lot Number", "Size", "Colour", "Quantity (Pieces)", "Remarks"]
    header_row = 7
    ws.row_dimensions[header_row].height = 24

    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=header_row, column=col_idx, value=header)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_thin

    current_row = 8
    total_qty = 0

    for idx, e in enumerate(entries, 1):
        ws.row_dimensions[current_row].height = 20
        try:
            d_str = datetime.strptime(e["entry_date"], "%Y-%m-%d").strftime("%d-%m-%Y")
        except:
            d_str = e["entry_date"]

        row_data = [idx, d_str, e["lot_number"], e["size"], e["colour"], e["quantity"], e["remarks"] or ""]

        for col_idx, val in enumerate(row_data, 1):
            cell = ws.cell(row=current_row, column=col_idx, value=val)
            cell.font = font_normal
            cell.border = border_thin
            
            if col_idx in [1, 2, 4]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx == 6:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = '#,##0'
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

            if current_row % 2 == 0:
                cell.fill = fill_zebra

        total_qty += e["quantity"]
        current_row += 1

    ws.row_dimensions[current_row].height = 22
    ws.cell(row=current_row, column=5, value="GRAND TOTAL:").font = font_bold
    ws.cell(row=current_row, column=5).alignment = Alignment(horizontal="right", vertical="center")
    
    total_cell = ws.cell(row=current_row, column=6, value=total_qty)
    total_cell.font = font_bold
    total_cell.number_format = '#,##0'
    total_cell.alignment = Alignment(horizontal="right", vertical="center")

    for col_idx in range(1, 8):
        ws.cell(row=current_row, column=col_idx).border = border_total

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    safe_party_name = "".join(c if c.isalnum() else "_" for c in party_name).upper()
    filename = f"{safe_party_name}_MANUAL_EXPORT.xlsx" if batch_num == 999 else f"{safe_party_name}_BATCH_{batch_num:03d}.xlsx"
    file_path = os.path.join(EXPORTS_DIR, filename)
    wb.save(file_path)
    return file_path, filename, total_qty

@app.post("/api/stock-entries")
def create_stock_entry(entry: StockEntryCreate):
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("SELECT name FROM parties WHERE id = ?", (entry.party_id,))
    party = cursor.fetchone()
    if not party:
        conn.close()
        raise HTTPException(status_code=404, detail="Party nahi mili!")
    party_name = party["name"]

    cursor.execute("SELECT * FROM stock_entries WHERE party_id = ? AND is_exported = 0 ORDER BY id ASC", (entry.party_id,))
    unexported_rows = cursor.fetchall()
    unexported_before = len(unexported_rows)
    
    if unexported_before >= 10:
        conn.close()
        raise HTTPException(status_code=400, detail="Is party ke 10 entries poore ho chuke hain! Pehle Excel download karein.")

    cursor.execute("SELECT COALESCE(MAX(batch_number), 0) as max_b FROM excel_batches WHERE party_id = ?", (entry.party_id,))
    current_batch_num = (cursor.fetchone()["max_b"] or 0) + 1

    cursor.execute("""
        INSERT INTO stock_entries (party_id, entry_date, lot_number, size, colour, quantity, remarks, batch_number, is_exported)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, 0)
    """, (
        entry.party_id, entry.entry_date, entry.lot_number.strip(), entry.size,
        entry.colour.strip(), entry.quantity, entry.remarks.strip() if entry.remarks else "", current_batch_num
    ))
    conn.commit()
    
    cursor.execute("SELECT * FROM stock_entries WHERE party_id = ? AND is_exported = 0 ORDER BY id ASC", (entry.party_id,))
    updated_unexported = cursor.fetchall()
    unexported_now = len(updated_unexported)
    
    batch_ready = False
    batch_id = None

    if unexported_now == 10:
        file_path, filename, total_pieces = generate_excel_batch_file(party_name, current_batch_num, [dict(r) for r in updated_unexported])
        
        cursor.execute("""
            INSERT INTO excel_batches (party_id, batch_number, file_path, total_entries, total_pieces)
            VALUES (?, ?, ?, 10, ?)
        """, (entry.party_id, current_batch_num, file_path, total_pieces))
        
        batch_id = cursor.lastrowid

        for r in updated_unexported:
            cursor.execute("UPDATE stock_entries SET is_exported = 1 WHERE id = ?", (r["id"],))
        conn.commit()
        batch_ready = True

    conn.close()
    
    return {
        "status": "success",
        "message": "Stock entry successfully save ho gayi!",
        "unexported_count": 0 if batch_ready else unexported_now,
        "batch_progress": f"{0 if batch_ready else unexported_now}/10",
        "batch_ready": batch_ready,
        "batch_number": current_batch_num,
        "batch_id": batch_id
    }

@app.put("/api/stock-entries/{entry_id}")
def update_stock_entry(entry_id: int, entry: StockEntryUpdate):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT is_exported FROM stock_entries WHERE id = ?", (entry_id,))
    row = cursor.fetchone()
    if not row:
        conn.close()
        raise HTTPException(status_code=404, detail="Entry nahi mili!")
    if row["is_exported"] == 1:
        conn.close()
        raise HTTPException(status_code=400, detail="Export ho chuke batch ki entry edit nahi ho sakti!")
        
    cursor.execute("""
        UPDATE stock_entries 
        SET entry_date = ?, lot_number = ?, size = ?, colour = ?, quantity = ?, remarks = ?
        WHERE id = ?
    """, (entry.entry_date, entry.lot_number.strip(), entry.size, entry.colour.strip(), entry.quantity, entry.remarks.strip() if entry.remarks else "", entry_id))
    conn.commit()
    conn.close()
    return {"status": "success", "message": "Entry successfully update ho gayi!"}

@app.delete("/api/stock-entries/{entry_id}")
def delete_stock_entry(entry_id: int):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT is_exported FROM stock_entries WHERE id = ?", (entry_id,))
    row = cursor.fetchone()
    if not row:
        conn.close()
        raise HTTPException(status_code=404, detail="Entry nahi mili!")
    if row["is_exported"] == 1:
        conn.close()
        raise HTTPException(status_code=400, detail="Export ho chuke batch ki entry delete nahi ki ja sakti!")
        
    cursor.execute("DELETE FROM stock_entries WHERE id = ?", (entry_id,))
    conn.commit()
    conn.close()
    return {"status": "success", "message": "Entry delete ho gayi aur batch counter update ho gaya!"}

@app.get("/api/excel-batches")
def get_excel_batches():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT eb.id, eb.batch_number, eb.total_entries, eb.total_pieces, eb.created_at, eb.file_path,
               p.name as party_name
        FROM excel_batches eb
        JOIN parties p ON eb.party_id = p.id
        ORDER BY eb.created_at DESC
    """)
    rows = cursor.fetchall()
    conn.close()
    
    batches = []
    for r in rows:
        filename = os.path.basename(r["file_path"])
        batches.append({
            "id": r["id"],
            "party_name": r["party_name"],
            "batch_number": r["batch_number"],
            "total_entries": r["total_entries"],
            "total_pieces": r["total_pieces"],
            "created_at": r["created_at"],
            "filename": filename
        })
    return batches

@app.get("/api/download-excel/{batch_id}")
def download_excel_file(batch_id: int):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT file_path FROM excel_batches WHERE id = ?", (batch_id,))
    row = cursor.fetchone()
    conn.close()
    
    if not row or not os.path.exists(row["file_path"]):
        raise HTTPException(status_code=404, detail="Excel file nahi mili!")
        
    file_path = row["file_path"]
    filename = os.path.basename(file_path)
    return FileResponse(file_path, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=filename)

@app.get("/api/parties/{party_id}/manual-export")
def manual_export_party(party_id: int):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM parties WHERE id = ?", (party_id,))
    party = cursor.fetchone()
    if not party:
        conn.close()
        raise HTTPException(status_code=404, detail="Party nahi mili!")
    
    cursor.execute("SELECT * FROM stock_entries WHERE party_id = ? ORDER BY entry_date DESC", (party_id,))
    entries = [dict(row) for row in cursor.fetchall()]
    conn.close()

    if not entries:
        raise HTTPException(status_code=400, detail="Is party ki koi stock entry nahi hai!")

    file_path, filename, total_pieces = generate_excel_batch_file(party["name"], 999, entries)
    return FileResponse(file_path, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=filename)

@app.get("/api/backup-database")
def backup_database():
    if not os.path.exists(DB_PATH):
        raise HTTPException(status_code=404, detail="Database file nahi mili!")
    return FileResponse(DB_PATH, media_type="application/x-sqlite3", filename="avkg_factory_backup.db")