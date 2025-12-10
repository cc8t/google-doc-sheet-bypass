import logging
import re
import requests

from bs4 import BeautifulSoup
from html4docx import HtmlToDocx
from io import BytesIO
import zipfile

from fastapi import FastAPI, Request
from fastapi.responses import StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates


app = FastAPI()
templates = Jinja2Templates(directory="templates")
app.mount("/static", StaticFiles(directory="static"), name="static")


def fetch_url(url: str) -> requests.Response | None:
    try:
        response = requests.get(url, timeout=60)
        return response if response.status_code == 200 else None
    except Exception:
        logging.exception(f"Failed to fetch: {url}")
        return None


def fetch_html(url: str) -> BeautifulSoup | None:
    response = fetch_url(url)
    if not response:
        return None
    return BeautifulSoup(response.text, "html.parser")


def get_sheet_ids(spreadsheet_id: str) -> list[str]:
    response = fetch_url(
        f"https://docs.google.com/spreadsheets/u/0/d/{spreadsheet_id}/htmlview"
    )
    if not response:
        return ["0"]

    sheet_ids = re.findall(r"gid=(\d+)", response.text)
    if not sheet_ids:
        return ["0"]

    return list(dict.fromkeys(sheet_ids))


def get_spreadsheet_title(spreadsheet_id: str) -> str:
    soup = fetch_html(
        f"https://docs.google.com/spreadsheets/u/0/d/{spreadsheet_id}/htmlview"
    )
    if not soup or not soup.title:
        return f"spreadsheet_{spreadsheet_id}"

    title = soup.title.string
    return re.sub(r"[./ ]+", "_", title)


def create_sheet(spreadsheet_id: str, sheet_id: str) -> BytesIO | None:
    urls = [
        f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=csv&gid={sheet_id}",
        f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/gviz/tq?tqx=out:csv&gid={sheet_id}",
    ]

    for url in urls:
        response = fetch_url(url)
        if not response:
            continue
        if response.content and not response.content.startswith(b"<!DOCTYPE"):
            buffer = BytesIO(response.content)
            buffer.seek(0)
            return buffer

    return None


def create_spreadsheet(spreadsheet_id: str) -> list[tuple[str, BytesIO]]:
    sheet_ids = get_sheet_ids(spreadsheet_id)
    files = []

    for sheet_id in sheet_ids:
        buffer = create_sheet(spreadsheet_id, sheet_id)
        if not buffer:
            logging.error(f"Failed to fetch sheet {sheet_id} for {spreadsheet_id}")
            continue

        filename = f"{sheet_id}.csv"
        files.append((filename, buffer))
        logging.info(f"Successfully fetched sheet {sheet_id}")

    if not files:
        raise Exception("No sheet could be fetched")

    return files


def create_document(doc_id: str) -> tuple[str, BytesIO]:
    soup = fetch_html(f"https://docs.google.com/document/d/{doc_id}/mobilebasic")

    title = f"document_{doc_id}"
    if soup and soup.title and soup.title.string:
        title = soup.title.string
        title = re.sub(r"[./ ]+", "_", title)

    parser = HtmlToDocx()
    outer = soup.find("div", {"class": "doc"})
    inner = outer.find("div")
    html = inner.prettify()
    doc = parser.parse_html_string(html)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return f"{title}.docx", buffer


@app.get("/")
def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/download")
async def download(request: Request):
    form = await request.form()
    doc_type = form.get("doc_type")
    doc_ids = form.getlist("doc_ids[]")

    zip_buffer = BytesIO()

    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        for doc_id in doc_ids:
            try:
                if doc_type == "docx":
                    filename, buffer = create_document(doc_id)
                    zip_file.writestr(filename, buffer.getvalue())

                elif doc_type == "csv":
                    dirname = get_spreadsheet_title(doc_id)
                    for filename, buffer in create_spreadsheet(doc_id):
                        zip_file.writestr(f"{dirname}/{filename}", buffer.getvalue())

                else:
                    raise ValueError("Invalid document type")

            except Exception:
                logging.exception(f"Failed to process document ID: {doc_id}")

    zip_buffer.seek(0)

    return StreamingResponse(
        zip_buffer,
        media_type="application/zip",
        headers={
            "Content-Disposition": "attachment; filename=google-doc-sheet-bypass.zip"
        },
    )
