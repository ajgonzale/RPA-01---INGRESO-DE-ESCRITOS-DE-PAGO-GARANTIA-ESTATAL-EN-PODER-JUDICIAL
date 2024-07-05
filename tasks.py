from robocorp import browser
from robocorp.tasks import task
from RPA.Excel.Files import Files as Excel

from pathlib import Path
import os
import requests

FILE_NAME = "challenge.xlsx"
OUTPUT_DIR = Path(os.environ.get('ROBOT_ARTIFACTS'))
EXCEL_URL = f"https://rpachallenge.com/assets/downloadFiles/{FILE_NAME}"


@task
def solve_challenge():
    """
    Solve the RPA challenge
    
    Downloads the source data excel and uses Playwright to solve rpachallenge.com from challenge
    """
    browser.configure(
        browser_engine="chromium",
        screenshot="only-on-failure",
        headless=False,
    )
    try:
        page = browser.goto("https://ojv.pjud.cl/kpitec-ojv-web/views/login.html")
        page.get_by_role("heading", name="Ingreso de demandas y escritos").wait_for()
        page.get_by_role("button", name="Clave Poder Judicial").first.click()
        page.get_by_role("textbox", name="Ej: 12345678 (Rut sin guión").click()
        page.get_by_role("textbox", name="Ej: 12345678 (Rut sin guión").fill("11346197123")
        page.locator("#inputPassword2C").click()
        page.locator("#inputPassword2C").fill("Talaveras1551+")
        page.get_by_role("button", name="Ingresar").click()
        page.locator("#roles-modal").get_by_text("Seleccione Perfil").wait_for()
        page.get_by_text("11.346.197-7").click()
        page.get_by_text("Ingresar Escrito").click()
        page.locator("#s2id_autogen1").get_by_role("link", name="Corte Suprema").click()
        page.get_by_role("option", name="Civil").click() 

        page.locator("label").filter(has_text="Fijar Datos").locator("path").nth(1).click()

        page.locator("xpath=//label[contains(.,'Tipo')]").wait_for()
        page.locator("xpath=//label[contains(.,'Tipo')]").click()
        page.press("body", "Tab")
        page.locator("xpath=//select[contains(@data-bind,'tiposCausas')]").select_option("option", label="C")
                        
        page.locator("xpath=//label[contains(.,'Tribunal')]").wait_for()
        page.locator("xpath=//label[contains(.,'Tribunal')]").click()
        page.press("body", "Tab") 
        page.press("body", "ArrowDown") 
        page.press("body", "1")      
        element = page.get_by_role("option", name="1º Juzgado De Letras de Angol").click()

        page.locator("xpath=//label[contains(.,'Rol')]/../input").fill("12345")
        page.press("body", "Tab")
        page.locator("xpath=//button[contains(.,' Consulta Rol')]").click()        

        page.locator(".toast-message").wait_for()
        page.press("body", "Tab")
        
        browser.screenshot(element)
        
        """
        excel_file = download_file(EXCEL_URL, OUTPUT_DIR, FILE_NAME)
        excel = Excel()
        excel.open_workbook(excel_file)
        rows = excel.read_worksheet_as_table("Sheet1", header=True)

        page = browser.goto("https://ojv.pjud.cl/kpitec-ojv-web/views/login.html")
        page.click("button:text('Start')")
        for row in rows:
            fill_and_submit_form(row)
        element = page.locator("css=div.congratulations")
        browser.screenshot(element)
        """
    finally:
        # Place for teardown and cleanups
        # Playwright handles browser closing
        print('Done')


def download_file(url: str, target_dir: Path, target_filename:str) -> str:
    """ 
    Downloads a file from the given url into the given folder with given filename. 
    """
    target_dir.mkdir(exist_ok=True)
    response = requests.get(url)
    response.raise_for_status()  # This will raise an exception if the request fails
    local_filename = Path(target_dir, target_filename)
    with open(local_filename, 'wb') as f:
        f.write(response.content)  # Write the content of the response to a file
    return local_filename


def fill_and_submit_form(row):
    """
    Fills a single form with the information of a single row in the Excel
    """
    page = browser.page()
    page.fill("//input[@ng-reflect-name='labelFirstName']", str(row["First Name"]))
    page.fill("//input[@ng-reflect-name='labelLastName']", str(row["Last Name"]))
    page.fill("//input[@ng-reflect-name='labelCompanyName']", str(row["Company Name"]))
    page.fill("//input[@ng-reflect-name='labelRole']", str(row["Role in Company"]))
    page.fill("//input[@ng-reflect-name='labelAddress']", str(row["Address"]))
    page.fill("//input[@ng-reflect-name='labelEmail']", str(row["Email"]))
    page.fill("//input[@ng-reflect-name='labelPhone']", str(row["Phone Number"]))
    page.click("input:text('Submit')")
