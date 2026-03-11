"""
MaisOLT Scraper
Uso: pip install selenium webdriver-manager pandas openpyxl
     python maisolt_scraper.py
"""

import sys, time, os
from datetime import datetime
from urllib.parse import urlparse, parse_qs

import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

BASE_URL = "https://cyberfly.maisolt.com.br"
OLT_IDS  = [i for i in range(1, 16) if i not in (7, 10, 14)]

STATUS_MAP = {
    "online":      "ONUs Online",
    "loss":        "ONUs LOSS",
    "sem energia": "ONUs Sem Energia",
    "inativo":     "ONUs Inativo",
}

COLUMNS = ["ID", "Nome da OLT", "ONUs Total", "ONUs Online",
           "ONUs LOSS", "ONUs Sem Energia", "ONUs Inativo"]


def criar_driver():
    opts = Options()
    opts.add_argument("--window-size=1280,900")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    service = Service(ChromeDriverManager().install())
    driver  = webdriver.Chrome(service=service, options=opts)
    driver.execute_script("Object.defineProperty(navigator,'webdriver',{get:()=>undefined})")
    return driver


def login(driver, email, password):
    print("[*] Fazendo login...")
    driver.get(BASE_URL)
    time.sleep(5)

    for _ in range(5):
        tipos = [i.get_attribute("type") for i in driver.find_elements(By.TAG_NAME, "input")]
        if "password" in tipos:
            break
        driver.get(f"{BASE_URL}/acesso_negado")
        time.sleep(5)
    else:
        input("[-] Formulário não encontrado. Faça login manualmente e pressione ENTER...")
        return True

    try:
        for inp in driver.find_elements(By.TAG_NAME, "input"):
            tipo = inp.get_attribute("type") or ""
            name = inp.get_attribute("name") or ""
            if tipo in ("text", "email") or name.lower() in ("username", "email", "user"):
                inp.clear(); inp.send_keys(email)
                break

        pwd = driver.find_element(By.CSS_SELECTOR, "input[type='password']")
        pwd.clear(); pwd.send_keys(password)
        driver.execute_script("document.querySelector('form').submit()")
        time.sleep(5)

        url = driver.current_url
        if all(x not in url for x in ["acesso_negado", "login", "after_login"]):
            print("[+] Login bem-sucedido!\n")
            return True

    except Exception as e:
        print(f"[-] Erro: {e}")

    input("[-] Login falhou. Faça login manualmente e pressione ENTER...")
    return True


def scrape_olt(driver, olt_id):
    driver.get(f"{BASE_URL}/olt/editar/{olt_id}")
    time.sleep(5)

    if any(x in driver.current_url for x in ["login", "acesso_negado", "after_login"]):
        print("sem acesso")
        return None

    try:
        WebDriverWait(driver, 25).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "a.ui.statistic.segment, a[class*='statistic']")
            )
        )
        time.sleep(2)
    except TimeoutException:
        print("sem dados")
        return None

    record = {col: None for col in COLUMNS}
    record["ID"] = olt_id

    for bloco in driver.find_elements(
        By.CSS_SELECTOR, "a.ui.statistic.segment, a[class*='statistic'][class*='segment']"
    ):
        href = bloco.get_attribute("href") or ""
        try:
            qs = parse_qs(urlparse(href).query)
        except Exception:
            continue

        olt_nome = qs.get("olt_nome", [None])[0]
        status   = qs.get("status",   [None])[0]

        number = None
        for sel in [".value", ".number", "[class*='value']"]:
            try:
                raw = bloco.find_element(By.CSS_SELECTOR, sel).text.strip()
                if raw.isdigit():
                    number = int(raw); break
            except Exception:
                pass
        if number is None:
            raw = bloco.text.strip().split("\n")[0]
            if raw.isdigit():
                number = int(raw)

        if olt_nome and not status:
            record["Nome da OLT"] = olt_nome
            record["ONUs Total"]  = number
        elif status:
            key = STATUS_MAP.get(status.lower())
            if key:
                record[key] = number

    if not record.get("Nome da OLT"):
        print("sem dados")
        return None

    return record


def export_excel(records, filename):
    df = pd.DataFrame(records, columns=COLUMNS)
    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="OLTs", index=False)
        wb, ws = writer.book, writer.sheets["OLTs"]

        hf  = Font(name="Arial", bold=True, color="FFFFFF", size=11)
        hb  = PatternFill("solid", start_color="1F4E79")
        ha  = Alignment(horizontal="center", vertical="center", wrap_text=True)
        df_ = Font(name="Arial", size=10)
        ab  = PatternFill("solid", start_color="DCE6F1")
        ca  = Alignment(horizontal="center", vertical="center")
        bd  = Border(left=Side(style="thin", color="B8CCE4"), right=Side(style="thin", color="B8CCE4"),
                     top=Side(style="thin", color="B8CCE4"), bottom=Side(style="thin", color="B8CCE4"))

        for c in range(1, len(COLUMNS)+1):
            cell = ws.cell(1, c)
            cell.font=hf; cell.fill=hb; cell.alignment=ha; cell.border=bd

        for r in range(2, len(df)+2):
            for c in range(1, len(COLUMNS)+1):
                cell = ws.cell(r, c)
                cell.font=df_; cell.alignment=ca; cell.border=bd
                if r % 2 == 0: cell.fill=ab

        for i, w in enumerate([6, 38, 14, 14, 12, 18, 16], 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"

        tr = len(df)+2
        ws.cell(tr, 1, "TOTAL").font = Font(name="Arial", bold=True)
        for ci in range(3, len(COLUMNS)+1):
            cl = get_column_letter(ci)
            cell = ws.cell(tr, ci, f"=SUM({cl}2:{cl}{tr-1})")
            cell.font=Font(name="Arial", bold=True, color="1F4E79")
            cell.alignment=ca; cell.border=bd
            cell.fill=PatternFill("solid", start_color="BDD7EE")

        wi = wb.create_sheet("Info")
        wi["A1"] = "Relatório MaisOLT — ONUs por OLT"
        wi["A1"].font = Font(name="Arial", bold=True, size=13, color="1F4E79")
        for i, (k, v) in enumerate([
            ("Data:",  datetime.now().strftime("%d/%m/%Y %H:%M:%S")),
            ("OLTs:",  len(records)),
            ("Fonte:", BASE_URL),
        ], 3):
            wi.cell(i, 1, k).font = Font(bold=True)
            wi.cell(i, 2, v)
        wi.column_dimensions["A"].width = 20
        wi.column_dimensions["B"].width = 35


def main():
    print("="*50)
    print("  MaisOLT Scraper — cyberfly.maisolt.com.br")
    print("="*50)

    email    = input("\nE-mail: ").strip()
    password = input("Senha:  ").strip()

    driver = criar_driver()
    try:
        if not login(driver, email, password):
            sys.exit(1)

        records = []
        for olt_id in OLT_IDS:
            print(f"  → OLT {olt_id:<3} ... ", end="", flush=True)
            data = scrape_olt(driver, olt_id)
            if data:
                records.append(data)
                print(f"✓  {data['Nome da OLT']}  "
                      f"[Online={data['ONUs Online']} | "
                      f"LOSS={data['ONUs LOSS']} | "
                      f"SE={data['ONUs Sem Energia']} | "
                      f"Inativo={data['ONUs Inativo']}]")
            else:
                print("— ignorado")

        if not records:
            print("\n[-] Nenhuma OLT encontrada.")
            sys.exit(1)

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        fn = f"maisolt_onus_{ts}.xlsx"
        export_excel(records, fn)
        print(f"\n✅ Concluído!  →  {os.path.abspath(fn)}")

    finally:
        driver.quit()


if __name__ == "__main__":
    main()
