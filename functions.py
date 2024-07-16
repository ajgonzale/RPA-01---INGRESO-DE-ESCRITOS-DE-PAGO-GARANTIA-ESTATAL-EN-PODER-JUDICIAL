from RPA.Excel.Files import Files as Excel
from robocorp import browser
from robocorp import windows
import time

def read_excel_preingreso(path: str):
    """
    open and read an excel file for preingreso

    """
    print('Preingreso')
    excel = Excel()
    excel.open_workbook(path)
    rows = excel.read_worksheet_as_table("Hoja1", header=True)
    numero = 2;
    preingresoIndex = 1;
    if not rows.columns.__contains__("PREINGRESO") :
        excel.insert_columns_before(1,1)
        #preingresoIndex = rows._add_column("PREINGRESO")
        excel.set_cell_value(1,1,"PREINGRESO")
        excel.save_workbook(path)
    else :
        preingresoIndex = rows.columns.index("PREINGRESO") + 1
    
    rows = excel.read_worksheet_as_table("Hoja1", header=True)
    print(preingresoIndex)
    for row in rows:
        if (str(row["PREINGRESO"]) != 'OK') :
            try :
               fill_form_escritos(row)
               excel.set_cell_value(numero,preingresoIndex,"OK")
               print(row["ROL"])
            except :
               excel.set_cell_value(numero,preingresoIndex,"ERROR")
               page = browser.page()
               page.get_by_text("Trámite Fácil").click()
               pass
            
        excel.save_workbook(path)
        numero = numero + 1
   
    excel.save_workbook(path)
    excel.close_workbook()
    print('Done Preingreso')

def fill_form_escritos(row):
    
    page = browser.page()
    page.get_by_text("Ingresar Escrito").click()

    time.sleep(1)

    page.get_by_text("Competencia").wait_for()
    page.get_by_text("Competencia").click()
    page.press("body", "Tab")
    page.keyboard.type("Civil") 
    page.press("body", "Tab")

    page.locator("label").filter(has_text="Fijar Datos").locator("path").nth(1).click()

    page.locator("xpath=//label[contains(.,'Tipo')]").wait_for()
    page.locator("xpath=//label[contains(.,'Tipo')]").click()
    page.press("body", "Tab")
    page.locator("xpath=//select[contains(@data-bind,'tiposCausas')]").select_option("option", label="C")
                        
    page.locator("xpath=//label[contains(.,'Tribunal')]").wait_for()
    page.locator("xpath=//label[contains(.,'Tribunal')]").click()
    page.press("body", "Tab") 
    page.press("body", "ArrowDown") 

    tribunal = str(row["TRIBUNAL"])
    tribunal = tribunal[0:len(tribunal)-1] + "º juzgado civil de santiago"
    page.keyboard.type(tribunal) 

    page.press("body", "Tab") 

    rol = str(row["ROL"])
    
    page.locator("xpath=//label[contains(.,'Rol')]/../input").fill(rol.split('-')[0])
    page.press("body", "Tab")
    page.press("body", "Tab")
    ano = rol.split('-')[1]
    page.keyboard.type(ano)  

    time.sleep(1)
    page.locator("xpath=//button[contains(.,' Consulta Rol')]").click()        

    page.locator("xpath=//label[contains(.,'Caratulado')]").wait_for()
    page.locator("xpath=//label[contains(.,'Caratulado')]").click()

    page.press("body", "Tab")
    """ Selecciono el cuaderno"""
    page.press("body", "ArrowDown")
    page.press("body", "Tab")

    time.sleep(1)

    page.keyboard.type("joa") 
    
    page.press("body", "Tab")
    page.press("body", "Tab")
    page.press("body", "Tab")
    
    page.keyboard.type("gen") 
     
    page.press("body", "Tab")
    
    time.sleep(1)

    page.keyboard.type("da cuenta de pag") 
    
    page.press("body", "Tab") 
  
    page.locator("xpath=//button[contains(.,' Grabar Escrito')]").click()

    time.sleep(1)

    page.get_by_text("Adjuntar Archivos").wait_for();
    page.get_by_text("Adjuntar Archivos").click()
    page.locator("#dDPrincipal").get_by_role("button", name="Adjuntar").click()
    app = windows.find_window('regex:.*Oficina Judicial Virtual - Portal - Chromium')
    app.find('control:"EditControl" and name:"File name:"').set_value("C:\Applications\Robocorp\RPA 01 - INGRESO DE ESCRITOS DE PAGO GARANTIA ESTATAL EN PODER JUDICIAL\input\Trib_1°_ROL_1452-2024_15.592.220-6.pdf")

    app.find('class:"Button" and name:"Open"').click();
    time.sleep(2)
    
    page.get_by_text("Escrito", exact=True).wait_for()
    page.get_by_text("Escrito", exact=True).click()
    page.get_by_role("button", name="Cerrar y Continuar").click()
    
    page.get_by_text("Trámite Fácil").click()


def read_excel_envio(path: str):
    """
    open and read an excel file for envio

    """
    print('Envio')
    excel = Excel()
    excel.open_workbook(path)
    rows = excel.read_worksheet_as_table("Hoja1", header=True)
    numero = 2;
    envioIndex = 1;
    if not rows.columns.__contains__("ENVIO") :
        excel.insert_columns_before(1,1)
        #envioIndex = rows._add_column("ENVIO")        
        excel.set_cell_value(1,1,"ENVIO")
        excel.save_workbook()
    else :
        envioIndex = rows.columns.index("ENVIO") + 1
    
    rows = excel.read_worksheet_as_table("Hoja1", header=True)
    print(envioIndex)
    for row in rows:
        if (str(row["ENVIO"]) != 'OK') :
            try :
                fill_form_envio(row)
                excel.set_cell_value(numero,envioIndex,"OK")
                print(row["ROL"])
            except :
                excel.set_cell_value(numero,len(row.keys()),"ERROR")
                page = browser.page()
                page.get_by_text("Mantenedor").click()
                pass
            
        excel.save_workbook()
        numero = numero + 1
   
    excel.save_workbook(path)
    excel.close_workbook()
    print('Done Envio')

def fill_form_envio(row):

    page = browser.page()
    page.get_by_text("Bandeja Escrito").click()
    
    time.sleep(2)

    page.locator("xpath=//label[contains(.,'Competencia:')]").wait_for()
    page.locator("xpath=//label[contains(.,'Competencia:')]").click()
    page.press("body", "Tab")
    page.keyboard.type("Civil") 
    page.press("body", "Tab")
    time.sleep(1)
    page.press("body", "Tab")
    time.sleep(1)
    page.press("body", "Tab")
    time.sleep(1)
    page.press("body", "Tab")
    time.sleep(1)
    page.press("body", "Tab")
    time.sleep(1)
    page.press("body", "Space") 
    page.press("body", "Tab")
    page.keyboard.type("C") 
    page.press("body", "Tab")

    rol = str(row["ROL"])   
    page.keyboard.type(rol.split('-')[0])  
    page.press("body", "Tab")
    page.keyboard.type(rol.split('-')[1])  
    page.press("body", "Tab")

    page.get_by_role("button", name="Consultar Escritos").click()

    time.sleep(2)

    page.locator("xpath=//label[contains(.,'Seleccionar Todo')]").first.click()  

    page.locator("xpath=//button[contains(.,'Eliminar Escritos')]").click()

    time.sleep(1)
    page.locator("#modalConfirmarEliminar").get_by_text("¿Está Seguro de Eliminar Escritos Seleccionados?").wait_for()
    page.locator("xpath=//div[@id='modalConfirmarEliminar']/div/div/div[2]/div/div/div[2]/button").click()

    page.get_by_text("Mantenedor").click()