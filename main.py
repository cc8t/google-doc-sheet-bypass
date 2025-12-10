from fastapi import FastAPI, Request
from fastapi.responses import StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
import requests
from bs4 import BeautifulSoup
import pandas as pd
from docx import Document
from io import BytesIO
import zipfile

app = FastAPI()
templates = Jinja2Templates(directory="templates")
app.mount("/static", StaticFiles(directory="static"), name="static")


def get_sheet_gids(doc_id):
    """Get all sheet GIDs from Google Spreadsheet"""
    sheet_gids = []
    
    try:
        # Try to get GIDs from the htmlview page
        url = f"https://docs.google.com/spreadsheets/u/0/d/{doc_id}/htmlview"
        r = requests.get(url, timeout=30)
        
        if r.status_code == 200:
            import re
            # Find all GID numbers in the HTML
            gid_matches = re.findall(r'gid=(\d+)', r.text)
            if gid_matches:
                # Remove duplicates and convert to int, then sort
                unique_gids = sorted(list(set(int(gid) for gid in gid_matches)))
                sheet_gids = unique_gids
    except Exception as e:
        print(f"Failed to get sheet GIDs: {e}")
    
    # If no GIDs found, default to first sheet
    if not sheet_gids:
        sheet_gids = [0]
    
    return sheet_gids


def create_csv(doc_id):
    """Create CSV files for all sheets in the spreadsheet"""
    # Get document name
    try:
        name_url = f"https://docs.google.com/spreadsheets/u/0/d/{doc_id}/htmlview"
        name_r = requests.get(name_url, timeout=10)
        soup = BeautifulSoup(name_r.text, "html.parser")
        base_name = soup.title.string.replace(".", "").replace("/", "").replace(" ", "_")
    except:
        base_name = f"spreadsheet_{doc_id}"
    
    # Get all sheet GIDs
    sheet_gids = get_sheet_gids(doc_id)
    
    csv_files = []
    
    for i, sheet_gid in enumerate(sheet_gids):
        # Simple naming: Sheet1, Sheet2, Sheet3, etc.
        sheet_name = f"Sheet{i+1}"
        
        # Try multiple URL formats for each sheet
        urls_to_try = [
            f"https://docs.google.com/spreadsheets/d/{doc_id}/export?format=csv&gid={sheet_gid}",
            f"https://docs.google.com/spreadsheets/d/{doc_id}/gviz/tq?tqx=out:csv&gid={sheet_gid}"
        ]
        
        for url in urls_to_try:
            try:
                r = requests.get(url, timeout=30)
                
                if r.status_code != 200:
                    continue
                
                # Check if we got actual CSV data (not an error page)
                if len(r.content) > 0 and not r.content.startswith(b'<!DOCTYPE'):
                    csv_buffer = BytesIO()
                    csv_buffer.write(r.content)
                    csv_buffer.seek(0)
                    
                    # Create filename with simple sheet name
                    if len(sheet_gids) > 1:
                        filename = f"{base_name}_{sheet_name}.csv"
                    else:
                        filename = f"{base_name}.csv"
                    
                    csv_files.append((filename, csv_buffer))
                    print(f"Successfully fetched {sheet_name} (GID: {sheet_gid})")
                    break
                    
            except Exception as e:
                print(f"Failed to fetch {sheet_name} (GID: {sheet_gid}): {e}")
                continue
    
    if not csv_files:
        raise Exception("Could not fetch any spreadsheet data")
    
    return csv_files


def create_docx(doc_id):
    url = f"https://docs.google.com/document/d/{doc_id}/mobilebasic"
    r = requests.get(url)
    soup = BeautifulSoup(r.text, "html.parser")
    doc_name = (
        soup.title.string.replace(".", "").replace("/", "").replace(" ", "_") + ".docx"
    )

    content_div = soup.find("div", {"class": "doc"})
    if content_div:
        content = content_div.find("div")
        text = content.get_text(separator="\n") if content else ""
    else:
        text = ""

    doc = Document()
    for line in text.split("\n"):
        doc.add_paragraph(line)

    docx_buffer = BytesIO()
    doc.save(docx_buffer)
    docx_buffer.seek(0)

    return doc_name, docx_buffer


@app.get("/")
def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/download")
async def download(request: Request):
    form = await request.form()
    doc_type = form.get("doc_type")
    doc_ids = form.getlist("doc_id[]")

    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        for doc_id in doc_ids:
            if not doc_id.strip():
                continue
            try:
                if doc_type == "csv":
                    csv_files = create_csv(doc_id)
                    for name, buffer in csv_files:
                        zip_file.writestr(name, buffer.getvalue())
                elif doc_type == "docx":
                    name, buffer = create_docx(doc_id)
                    zip_file.writestr(name, buffer.getvalue())
                else:
                    raise ValueError("Invalid document type")
            except Exception as e:
                print(f"Failed to fetch {doc_id}: {e}")

    zip_buffer.seek(0)
    return StreamingResponse(
        zip_buffer,
        media_type="application/zip",
        headers={
            "Content-Disposition": "attachment; filename=google-doc-sheet-bypass.zip"
        },
    )
