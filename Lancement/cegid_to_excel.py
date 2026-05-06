import json
import os
import re
from datetime import datetime
from getpass import getpass
from pathlib import Path

from openpyxl import Workbook
from playwright.sync_api import sync_playwright

URL = "https://script.cegid.fr/"

BASE_DIR = Path(r"C:\Users\Ali LOUNNAS\Desktop\Automate\Script Cegid")
LAUNCH_DIR = BASE_DIR / "Lancement"
OUTPUT_DIR = BASE_DIR / "Extractions"
CONFIG_PATH = LAUNCH_DIR / "config.json"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
LAUNCH_DIR.mkdir(parents=True, exist_ok=True)


def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def make_excel_path():
    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    return OUTPUT_DIR / f"planning_import_{ts}.xlsx"


def clean_text(value):
    return re.sub(r"\s+", " ", (value or "")).strip()


def parse_business_text(text_raw):
    text = clean_text(text_raw)

    if "*" in text:
        left, client = text.split("*", 1)
    else:
        left, client = text, ""

    parts = [p.strip() for p in left.split(" - ")]

    produit = parts[0] if len(parts) > 0 else ""
    mission = parts[1] if len(parts) > 1 else ""
    type_projet = " - ".join(parts[2:]).strip() if len(parts) > 2 else ""

    return produit, mission, type_projet, client.strip()


def parse_hours(hours_raw):
    match = re.search(r"(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2})", hours_raw or "")
    if match:
        return match.group(1), match.group(2)
    return "", ""


def load_config():
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                return {
                    "email": data.get("email", ""),
                    "password": data.get("password", ""),
                }
        except Exception:
            return {"email": "", "password": ""}
    return {"email": "", "password": ""}


def save_config(email, password):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(
            {"email": email, "password": password},
            f,
            indent=2,
            ensure_ascii=False,
        )


def ask_email(current_value):
    if current_value:
        value = input(f"Email [{current_value}] : ").strip()
        return value if value else current_value
    return input("Email : ").strip()


def ask_password(current_value):
    if current_value:
        value = getpass("Mot de passe [laisser vide pour garder l'existant] : ").strip()
        return value if value else current_value
    return getpass("Mot de passe : ").strip()


def build_workbook():
    wb = Workbook()
    ws = wb.active
    ws.title = "Donnees"
    ws.append([
        "Date",
        "Produit",
        "Mission",
        "TypeProjet",
        "Client",
        "HeureDebut",
        "HeureFin",
        "Statut",
        "DateImport",
    ])
    return wb, ws


def first_visible(locator):
    count = locator.count()
    for i in range(count):
        item = locator.nth(i)
        try:
            if item.is_visible():
                return item
        except Exception:
            pass
    return None


def click_first_visible(locator):
    item = first_visible(locator)
    if item:
        item.click()
        return True
    return False


def fill_first_visible(locator, value):
    item = first_visible(locator)
    if item:
        item.fill(value)
        return True
    return False


def open_if_closed(header):
    try:
        icon = header.locator("i.glyphicon").first
        icon_class = icon.get_attribute("class") or ""

        # D'après ton HTML :
        # chevron-up = fermé
        # chevron-down = ouvert
        if "glyphicon-chevron-up" in icon_class:
            header.click()
            return
        if "glyphicon-chevron-down" in icon_class:
            return

        header.click()
    except Exception:
        header.click()


def get_panel_content(header):
    panel = header.locator("xpath=ancestor::div[contains(@class,'panel')][1]")
    content = panel.locator("div.panel-collapse").first

    if content.count() == 0:
        content = header.locator(
            "xpath=ancestor::div[contains(@class,'panel-heading')]/following-sibling::div[contains(@class,'panel-collapse')][1]"
        )
        if content.count() == 0:
            return None

    try:
        content.wait_for(state="attached", timeout=5000)
    except Exception:
        return None

    return content


def main():
    config = load_config()

    env_email = os.getenv("CEGID_EMAIL", "").strip()
    env_password = os.getenv("CEGID_PASSWORD", "").strip()

    email = env_email if env_email else ask_email(config.get("email", ""))
    password = env_password if env_password else ask_password(config.get("password", ""))

    save_config(email, password)

    excel_path = make_excel_path()
    wb, ws = build_workbook()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()
        page.set_default_timeout(30000)

        try:
            print("Ouverture du site...")
            page.goto(URL, wait_until="domcontentloaded")
            page.wait_for_timeout(2000)

            print("Recherche du premier bouton de connexion...")
            click_first_visible(
                page.get_by_role("button", name=re.compile(r"se connecter", re.I))
            )
            page.wait_for_timeout(2000)

            print("Saisie de l'email...")
            filled_email = False

            if page.locator("#userEmailEntryInput").count() > 0:
                filled_email = fill_first_visible(page.locator("#userEmailEntryInput"), email)

            if not filled_email and page.locator("#selfAssertedEmail").count() > 0:
                filled_email = fill_first_visible(page.locator("#selfAssertedEmail"), email)

            if not filled_email and page.locator('input[name="email"]').count() > 0:
                filled_email = fill_first_visible(page.locator('input[name="email"]'), email)

            if not filled_email and page.locator('input[type="email"]').count() > 0:
                filled_email = fill_first_visible(page.locator('input[type="email"]'), email)

            if not filled_email:
                textbox = first_visible(page.get_by_role("textbox"))
                if textbox:
                    textbox.fill(email)
                    filled_email = True

            if not filled_email:
                raise Exception("Champ email introuvable.")

            page.wait_for_timeout(800)

            print("Validation email...")
            if not click_first_visible(
                page.get_by_role("button", name=re.compile(r"se connecter|continuer", re.I))
            ):
                raise Exception("Bouton après email introuvable.")

            page.wait_for_timeout(2500)

            print("Saisie du mot de passe...")
            filled_password = False

            if page.locator('input[type="password"]').count() > 0:
                filled_password = fill_first_visible(page.locator('input[type="password"]'), password)

            if not filled_password:
                raise Exception("Champ mot de passe introuvable.")

            page.wait_for_timeout(800)

            print("Validation mot de passe...")
            if not click_first_visible(
                page.get_by_role("button", name=re.compile(r"continuer|se connecter", re.I))
            ):
                raise Exception("Bouton après mot de passe introuvable.")

            page.wait_for_timeout(5000)

            print("Attente du calendrier...")
            page.wait_for_selector("div.panel-heading a.accordion-toggle")

            headers = page.locator("div.panel-heading a.accordion-toggle")
            header_count = headers.count()
            print("Nombre de jours trouvés :", header_count)

            total_rows = 0

            for i in range(header_count):
                header = headers.nth(i)

                if not header.is_visible():
                    continue

                date_locator = header.locator("span")
                if date_locator.count() == 0:
                    continue

                date = clean_text(date_locator.first.inner_text())
                print(f"Date trouvée : {date}")

                open_if_closed(header)
                page.wait_for_timeout(1000)

                content = get_panel_content(header)
                if content is None:
                    print(f"Bloc introuvable pour {date}, jour ignoré.")
                    continue

                rows = content.locator("li.list-group-item")
                row_count = rows.count()
                print(f"Nombre de lignes pour {date} : {row_count}")

                for j in range(row_count):
                    row = rows.nth(j)

                    title_locator = row.locator('a[ng-href^="#/appointment/"]')
                    if title_locator.count() == 0:
                        continue

                    text = clean_text(title_locator.first.inner_text())
                    if not text:
                        continue

                    hours = ""
                    hours_locator = row.locator("label")
                    if hours_locator.count() > 0:
                        hours = clean_text(hours_locator.first.inner_text())

                    produit, mission, type_projet, client = parse_business_text(text)
                    h_debut, h_fin = parse_hours(hours)

                    statut = "Validé" if row.locator("span.glyphicon-ok").count() > 0 else "Non validé"

                    ws.append([
                        date,
                        produit,
                        mission,
                        type_projet,
                        client,
                        h_debut,
                        h_fin,
                        statut,
                        now_str(),
                    ])
                    total_rows += 1

            wb.save(excel_path)
            print(f"Nombre total de lignes exportées : {total_rows}")
            print(f"Excel généré : {excel_path.resolve()}")

            if total_rows == 0:
                print("Aucune ligne n'a été exportée. Vérifier les sélecteurs ou le chargement de la page.")

        finally:
            context.close()
            browser.close()


if __name__ == "__main__":
    main()
