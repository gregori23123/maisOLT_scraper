"""
MaisOLT Scraper — Paralelo
Uso: pip install selenium webdriver-manager openpyxl
     python maisolt_scraper.py
"""

import sys, time, os
from datetime import datetime
from urllib.parse import urlparse, parse_qs
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading

import openpyxl

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

BASE_URL   = "https://cyberfly.maisolt.com.br"
OLT_IDS    = [i for i in range(1, 19) if i not in (7, 10, 14)]
WORKERS    = 4   # número de navegadores paralelos
print_lock = threading.Lock()

STATUS_MAP = {
    "inativo":     "Inativo",
    "online":      "Online",
    "sem energia": "Power Fail",
    "loss":        "Loss",
}

COLUMNS = ["ID", "Nome da OLT", "Data", "Hora", "Inativo",
           "Online", "Power Fail", "Loss"]


def criar_driver(headless=False):
    opts = Options()
    opts.add_argument("--window-size=1280,900")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    if headless:
        opts.add_argument("--headless=new")
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
        print(f"[-] Erro no login: {e}")

    input("[-] Login falhou. Faça login manualmente e pressione ENTER...")
    return True


def scrape_olt(cookies, olt_id):
    """Cria um driver próprio, injeta cookies da sessão e raspa a OLT."""
    driver = criar_driver(headless=True)
    try:
        # injeta cookies para reaproveitar sessão sem novo login
        driver.get(BASE_URL)
        time.sleep(2)
        for cookie in cookies:
            try:
                driver.add_cookie(cookie)
            except Exception:
                pass

        driver.get(f"{BASE_URL}/olt/editar/{olt_id}")
        time.sleep(3)

        if any(x in driver.current_url for x in ["login", "acesso_negado", "after_login"]):
            return olt_id, None, "sem acesso"

        try:
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "a.ui.statistic.segment, a[class*='statistic']")
                )
            )
            time.sleep(1)
        except TimeoutException:
            return olt_id, None, "sem dados"

        record = {col: None for col in COLUMNS}
        record["ID"] = olt_id
        now = datetime.now()
        record["Data"] = now.strftime("%d/%m/%Y")
        record["Hora"] = now.strftime("%H:%M")

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
            elif status:
                key = STATUS_MAP.get(status.lower())
                if key:
                    record[key] = number

        if not record.get("Nome da OLT"):
            return olt_id, None, "sem dados"

        return olt_id, record, "ok"

    finally:
        driver.quit()


def export_xlsx(records, filename):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "OLTs"

    ws.append(COLUMNS)

    for r in sorted(records, key=lambda x: x["ID"]):
        row = [r.get(col) if r.get(col) is not None else 0 for col in COLUMNS]
        ws.append(row)



    wb.save(filename)


def main():
    print("="*50)
    print("  MaisOLT Scraper — cyberfly.maisolt.com.br")
    print(f"  Modo paralelo: {WORKERS} workers")
    print("="*50)

    email    = input("\nE-mail: ").strip()
    password = input("Senha:  ").strip()

    # login em um driver principal para obter cookies de sessão
    main_driver = criar_driver(headless=False)
    try:
        if not login(main_driver, email, password):
            sys.exit(1)
        cookies = main_driver.get_cookies()
    finally:
        main_driver.quit()

    print(f"[*] Raspando {len(OLT_IDS)} OLTs com {WORKERS} workers em paralelo...\n")

    records = []
    with ThreadPoolExecutor(max_workers=WORKERS) as executor:
        futures = {executor.submit(scrape_olt, cookies, olt_id): olt_id for olt_id in OLT_IDS}
        for future in as_completed(futures):
            olt_id, data, status = future.result()
            with print_lock:
                if data:
                    records.append(data)
                    print(f"  ✓ OLT {olt_id:<3} {data['Nome da OLT']:<35} "
                          f"Inativo={data['Inativo']} | "
                          f"Online={data['Online']} | "
                          f"PF={data['Power Fail']} | "
                          f"Loss={data['Loss']}")
                else:
                    print(f"  — OLT {olt_id:<3} {status}")

    if not records:
        print("\n[-] Nenhuma OLT encontrada.")
        sys.exit(1)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    fn = f"maisolt_onus_{ts}.xlsx"
    export_xlsx(records, fn)
    print(f"\n✅ Concluído!  →  {os.path.abspath(fn)}")


if __name__ == "__main__":
    main()
