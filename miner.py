from __future__ import annotations

import time
import os
import random
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager


# =========================
# CONFIG
# =========================
HEADLESS = False                   # True = no abre ventana de Chrome
INPUT_SHEET_NAME = None           # None = primera hoja
SLEEP_BETWEEN_PARTS = (3.5, 7.1)  # anti-throttle suave
PAGE_TIMEOUT = 45                 # timeout
SAVE_EVERY = 10                    # ✅ guarda cada 10 SKUs para no perder progreso

BASE_URL = "https://showmetheparts.com/standard-parts/standard-{part}.html"


# =========================
# UI Helpers
# =========================
def pick_input_file(title: str = "Selecciona el Excel con SKUs") -> Path:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = filedialog.askopenfilename(
        title=title,
        filetypes=[
            ("Excel", "*.xlsx *.xlsm *.xls"),
            ("Todos los archivos", "*.*"),
        ],
    )
    root.destroy()
    if not path:
        raise SystemExit("Cancelado: no seleccionaste archivo.")
    return Path(path)


def pick_output_dir(title: str = "Selecciona la carpeta de salida") -> Path:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    folder = filedialog.askdirectory(title=title)
    root.destroy()
    if not folder:
        raise SystemExit("Cancelado: no seleccionaste carpeta de salida.")
    out = Path(folder)
    out.mkdir(parents=True, exist_ok=True)
    return out


def info_popup(msg: str) -> None:
    print(msg)
    try:
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("Listo", msg)
        root.destroy()
    except Exception:
        pass


def timestamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


# =========================
# Helpers
# =========================
def jitter(a_b: Tuple[float, float]) -> None:
    a, b = a_b
    time.sleep(random.uniform(a, b))


def read_parts_from_excel(path: Path) -> List[str]:
    df = pd.read_excel(path, sheet_name=INPUT_SHEET_NAME)

    if isinstance(df, dict):
        df = next(iter(df.values()))

    cols_lower = {str(c).strip().lower(): c for c in df.columns}
    sku_col = cols_lower.get("sku")

    if sku_col is None:
        sku_col = df.columns[0]  # fallback

    parts = (
        df[sku_col]
        .dropna()
        .astype(str)
        .str.strip()
        .str.upper()
        .tolist()
    )

    clean: List[str] = []
    for p in parts:
        if not p or p.lower() in ("nan", "none"):
            continue
        if len(p) < 2:
            continue
        clean.append(p)
    return clean


def wait_page_ready(driver: webdriver.Chrome, timeout: int) -> None:
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )
    time.sleep(0.35)


def page_status_hint(driver: webdriver.Chrome) -> str:
    src = (driver.page_source or "").lower()

    nores_signals = [
        "page not found",
        "not found",
        "404",
        "no results",
        "no result",
        "nothing found",
        "could not find",
    ]
    block_signals = [
        "captcha",
        "verify you are",
        "access denied",
        "unusual traffic",
        "robot",
        "blocked",
        "cloudflare",
    ]

    if any(s in src for s in block_signals):
        return "BLOQUEADO"
    if any(s in src for s in nores_signals):
        return "NO_RESULTS"
    return "OK"


def save_data_excel(all_rows: List[Dict[str, Any]], output_xlsx: Path) -> int:
    df_data = pd.DataFrame(all_rows) if all_rows else pd.DataFrame(columns=["PartNumber"])
    if not df_data.empty:
        cols = ["PartNumber"] + [c for c in df_data.columns if c != "PartNumber"]
        df_data = df_data[cols]
    df_data.to_excel(output_xlsx, index=False)
    return len(df_data)


def save_links_excel(link_rows: List[Dict[str, Any]], links_xlsx: Path) -> int:
    # ✅ Excel pedido: Col A SKU, Col B LINK. (Le agrego STATUS y NOTE para que sepas qué pasó)
    df = pd.DataFrame(link_rows) if link_rows else pd.DataFrame(columns=["SKU", "LINK", "STATUS", "NOTE"])
    # Orden fijo
    cols = ["SKU", "LINK", "STATUS", "NOTE"]
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    df = df[cols]
    df.to_excel(links_xlsx, index=False)
    return len(df)


# =========================
# Selenium
# =========================
def build_driver(headless: bool = False) -> webdriver.Chrome:
    options = webdriver.ChromeOptions()

    # =========================
    # ✅ PERFIL PERSISTENTE (ANTI-CAPTCHA)
    # =========================
    profile_dir = os.path.join(
        str(Path.home()),
        "selenium_profile_showmetheparts"
    )
    options.add_argument(f"--user-data-dir={profile_dir}")
    options.add_argument("--profile-directory=Default")

    # =========================
    # HEADLESS
    # =========================
    if headless:
        options.add_argument("--headless=new")

    # =========================
    # OPCIONES BASE
    # =========================
    options.add_argument("--window-size=1400,900")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    # =========================
    # ANTI-DETECCIÓN BÁSICA
    # =========================
    options.add_argument("--lang=en-US")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    # =========================
    # QUITAR navigator.webdriver
    # =========================
    try:
        driver.execute_cdp_cmd(
            "Page.addScriptToEvaluateOnNewDocument",
            {
                "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
            }
        )
    except Exception:
        pass

    # =========================
    # TIMEOUT
    # =========================
    try:
        driver.set_page_load_timeout(PAGE_TIMEOUT)
    except Exception:
        pass

    return driver

def find_buyers_guide_table(driver: webdriver.Chrome) -> Optional[Any]:
    WebDriverWait(driver, PAGE_TIMEOUT).until(
        EC.presence_of_element_located((By.TAG_NAME, "body"))
    )

    xps = [
        "//h2[contains(translate(., 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), 'BUYERS GUIDE')]/following::table[1]",
        "//h3[contains(translate(., 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), 'BUYERS GUIDE')]/following::table[1]",
        "//*[contains(translate(., 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), 'BUYERS GUIDE')]/following::table[1]",
    ]
    for xp in xps:
        try:
            return driver.find_element(By.XPATH, xp)
        except Exception:
            pass

    # fallback: tabla con más filas
    try:
        WebDriverWait(driver, 6).until(EC.presence_of_element_located((By.TAG_NAME, "table")))
    except Exception:
        return None

    tables = driver.find_elements(By.TAG_NAME, "table")
    if not tables:
        return None

    best = None
    best_rows = -1
    for t in tables:
        try:
            n = len(t.find_elements(By.TAG_NAME, "tr"))
            if n > best_rows:
                best_rows = n
                best = t
        except Exception:
            continue

    return best


def extract_table_rows(table, part: str) -> List[Dict[str, Any]]:
    header_cells = table.find_elements(By.TAG_NAME, "th")
    headers = [h.text.strip() for h in header_cells if h.text and h.text.strip()]

    rows_out: List[Dict[str, Any]] = []

    for tr in table.find_elements(By.TAG_NAME, "tr"):
        cells = tr.find_elements(By.TAG_NAME, "td")
        if not cells:
            continue

        values = [c.text.strip() for c in cells]
        if not any(v for v in values):
            continue

        row_dict: Dict[str, Any] = {"PartNumber": part}

        if headers and len(headers) == len(values):
            for h, v in zip(headers, values):
                row_dict[h] = v
        elif headers:
            for i, h in enumerate(headers):
                if i < len(values):
                    row_dict[h] = values[i]
            if len(values) > len(headers):
                for j in range(len(headers), len(values)):
                    row_dict[f"Extra{j-len(headers)+1}"] = values[j]
        else:
            for i, v in enumerate(values, start=1):
                row_dict[f"Col{i}"] = v

        rows_out.append(row_dict)

    return rows_out


# =========================
# Main
# =========================
def main():
    print(">>> Iniciando buyers_guide_general.py (SIN carpetas debug + Excel links autosave cada 5)")

    input_xlsx = pick_input_file("Selecciona el Excel con SKUs (columna SKU o primera columna)")
    out_dir = pick_output_dir("Selecciona la carpeta donde guardar")

    parts = read_parts_from_excel(input_xlsx)
    if not parts:
        raise SystemExit("No hay SKUs válidos en el archivo.")

    output_data_xlsx = out_dir / f"{input_xlsx.stem}_buyers_guide_{timestamp()}.xlsx"
    output_links_xlsx = out_dir / f"{input_xlsx.stem}_links_log_{timestamp()}.xlsx"

    print(f"\n📄 Archivo entrada: {input_xlsx}")
    print(f"📌 SKUs a procesar: {len(parts)}")
    print(f"📄 Salida DATA:  {output_data_xlsx.name}")
    print(f"📄 Salida LINKS: {output_links_xlsx.name}\n")

    driver = build_driver(headless=HEADLESS)

    all_rows: List[Dict[str, Any]] = []
    link_rows: List[Dict[str, Any]] = []

    ok = no_table = blocked = err = 0

    try:
        for idx, part in enumerate(parts, start=1):
            url = BASE_URL.format(part=part)
            print(f"[{idx}/{len(parts)}] {part} -> {url}")

            try:
                driver.get(url)
                wait_page_ready(driver, PAGE_TIMEOUT)

                hint = page_status_hint(driver)
                if hint == "BLOQUEADO":
                    blocked += 1
                    print("   🚫 BLOQUEADO/CAPTCHA")
                    link_rows.append({"SKU": part, "LINK": url, "STATUS": "BLOQUEADO", "NOTE": "Captcha / antibot"})
                elif hint == "NO_RESULTS":
                    no_table += 1
                    print("   ❌ NO RESULTS / NOT FOUND")
                    link_rows.append({"SKU": part, "LINK": url, "STATUS": "SIN_DATOS", "NOTE": "No results/404"})
                else:
                    table = find_buyers_guide_table(driver)
                    if table is None:
                        no_table += 1
                        print("   ❌ No encontré tabla")
                        link_rows.append({"SKU": part, "LINK": url, "STATUS": "SIN_TABLA", "NOTE": "No table found"})
                    else:
                        rows = extract_table_rows(table, part)
                        if not rows:
                            no_table += 1
                            print("   ⚠️ Tabla pero sin filas")
                            link_rows.append({"SKU": part, "LINK": url, "STATUS": "TABLA_VACIA", "NOTE": "Table no rows"})
                        else:
                            all_rows.extend(rows)
                            ok += 1
                            print(f"   ✅ Filas agregadas: {len(rows)}")
                            link_rows.append({"SKU": part, "LINK": url, "STATUS": "OK", "NOTE": f"{len(rows)} filas"})

            except Exception as e:
                err += 1
                print(f"   ⚠️ ERROR: {e}")
                link_rows.append({"SKU": part, "LINK": url, "STATUS": "ERROR", "NOTE": str(e)[:240]})

            # ✅ Guardado incremental cada 5 SKUs procesados
            if SAVE_EVERY and (idx % SAVE_EVERY == 0):
                saved_data = save_data_excel(all_rows, output_data_xlsx)
                saved_links = save_links_excel(link_rows, output_links_xlsx)
                print(f"💾 Autosave: DATA={saved_data} filas | LINKS={saved_links} registros")

            jitter(SLEEP_BETWEEN_PARTS)

    finally:
        driver.quit()

    # Guardado final
    saved_data = save_data_excel(all_rows, output_data_xlsx)
    saved_links = save_links_excel(link_rows, output_links_xlsx)

    resumen = (
        f"✅ Listo\n\n"
        f"DATA:\n{output_data_xlsx}\n\n"
        f"LINKS LOG:\n{output_links_xlsx}\n\n"
        f"Resumen:\n"
        f"SKUs totales: {len(parts)}\n"
        f"OK (con datos): {ok}\n"
        f"Sin tabla/filas: {no_table}\n"
        f"Bloqueados: {blocked}\n"
        f"Errores: {err}\n"
        f"Filas DATA guardadas: {saved_data}\n"
        f"Registros LINKS guardados: {saved_links}\n"
    )
    print("\n" + resumen)
    info_popup(resumen)


if __name__ == "__main__":
    main()
