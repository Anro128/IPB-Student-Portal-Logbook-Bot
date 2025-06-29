import requests
from bs4 import BeautifulSoup
import pandas as pd
import os
import mimetypes
import browser_cookie3
from datetime import datetime

def getCookies(domain, browser = 'firefox', cookieName=''):
    Cookies = {}
    Bcookies = ""
    if(browser.lower() == 'firefox'):
        Bcookies = list(browser_cookie3.firefox())
    elif(browser.lower() == 'chrome'):
        Bcookies = list(browser_cookie3.chrome())

    for cookie in Bcookies:
        if domain in cookie.domain:
            Cookies[cookie.name] = cookie.value

    if cookieName != '':
        try:
            return Cookies[cookieName] 
        except KeyError:
            return {}  
    else:
        return Cookies 

def is_valid_date(tanggal_str):
    try:
        datetime.strptime(tanggal_str, "%d/%m/%Y")
        return True
    except ValueError:
        return False    

def is_valid_time(waktu_str):
    try:
        datetime.strptime(waktu_str, "%H:%M")
        return True
    except ValueError:
        return False

aktivitas_id = input("Aktivitas ID: ")
browser = input("Browsers used: ")

BASE_URL      = "https://studentportal.ipb.ac.id"
COOKIE_JAR    = getCookies("studentportal.ipb.ac.id", browser)
EXCEL_PATH    = "data.xlsx" 
SHEET_NAME    = "Sheet1"      
OUTPUT_LOG    = "result.csv" 

df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)

session = requests.Session()
session.cookies.update(COOKIE_JAR)
session.headers.update({
    "User-Agent":        "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
    "Accept":            "*/*",
    "X-Requested-With":  "XMLHttpRequest",
})

results = []

for idx, row in df.iterrows():
    try:
        if row["Tstart"] >= row['Tend']:
            raise ValueError(f"Start time {row['Tstart']} must be earlier than end time {row['Tend']}\nData {idx+1} skipped.")

        if not is_valid_date(row['Waktu']):
            raise ValueError(f"Invalid date format, use DD/MM/YYYY. \nData {idx+1} skipped.")
        
        if not is_valid_time(row['Tstart']):
            raise ValueError(f"Invalid time start format, use HH:MM. \nData {idx+1} skipped.")
        
        if not is_valid_time(row['Tend']):
            raise ValueError(f"Invalid time end format, use HH:MM. \nData {idx+1} skipped.")
        
        modal_url = (
            f"{BASE_URL}/Kegiatan/LogAktivitasKampusMerdeka/Tambah"
            f"?aktivitasId={aktivitas_id}"
        )
        resp = session.get(modal_url)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")

        hidden = {
            tag["name"]: tag.get("value", "")
            for tag in soup.select("form input[type=hidden]")
        }

        data = hidden.copy()
        data.update({
            "Waktu":str(row['Waktu']),
            "Tmw":str(row['Tstart']),
            "Tsw":str(row['Tend']),
            "JenisLogbookKegiatanKampusMerdekaId":str(row['JenisLogId']),
            "ListDosenPembimbing[0].Value":str("true").lower(),
            "IsLuring": "true" if row['IsLuring'] == 1 else "false" if row['IsLuring'] == 0 else "",
            "Lokasi":row['Lokasi'],
            "Keterangan":row['Keterangan'],
        })

        file_path = row['FilePath']
        filename  = os.path.basename(file_path)
        ctype, _  = mimetypes.guess_type(file_path)
        files = {
            "File": (
                filename,
                open(file_path, "rb"),
                ctype or "application/octet-stream"
            )
        }

        post_url = f"{BASE_URL}/Kegiatan/LogAktivitasKampusMerdeka/Tambah"
        r_post = session.post(post_url, data=data, files=files, allow_redirects=False)

        success = False
        if r_post.status_code == 302:
            list_url = (
                f"{BASE_URL}/Kegiatan/LogAktivitasKampusMerdeka/Index/"
                f"{aktivitas_id}"
            )
            r_list = session.get(list_url)
            success = row['Keterangan'] in r_list.text
        else:
            success = False

        results.append({
            "row": idx,
            "status_code": r_post.status_code,
        })
        print(f"data {idx+1} uploaded successfully")

    except Exception as e:
        results.append({
            "row": idx,
            "status_code": "ERROR",
            "error": str(e)
        })
        print(f"data {idx+1} failed to upload: {str(e)}")


out_df = pd.DataFrame(results)
out_df.to_csv(OUTPUT_LOG, index=False)
print(f"Finished. See {OUTPUT_LOG} for a summary of the results.")
