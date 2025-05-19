from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

driver_path = "C:/Users/Cobos/OneDrive/Documentos/chromedriver-win64/chromedriver.exe"
service = Service(driver_path)

options = Options()
options.add_argument("--start-maximized")
options.add_argument("--disable-notifications")
options.add_experimental_option("excludeSwitches", ["enable-automation"])

MOODLE_USER = "g.cobos.sumire@isur.edu.pe"
MOODLE_PASSWORD = "11_de_SETIEMBRE_2001"

driver = webdriver.Chrome(service=service, options=options)

try:
    print("\n=== INICIANDO AUTOMATIZACIÓN DE MOODLE ISUR ===")
    
    print("\n[1/6] Cargando página de login...")
    driver.get("https://login.microsoftonline.com/isur.edu.pe/oauth2/authorize?response_type=code&client_id=67d06b8c-6885-4dde-8470-abe8ebe9b279&scope=openid%20profile%20email&nonce=N682aa61624ed0&response_mode=form_post&state=dzk5rVm6DyCTuyX&redirect_uri=https%3A%2F%2Fmoodle.isur.edu.pe%2Fauth%2Foidc%2F&resource=https%3A%2F%2Fgraph.microsoft.com&sso_reload=true")
    time.sleep(3)

    print("[2/6] Ingresando usuario...")
    email_field = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='loginfmt']"))
    )
    email_field.send_keys(MOODLE_USER)
    
    next_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "idSIButton9"))
    )
    next_button.click()
    time.sleep(3)

    print("[3/6] Ingresando contraseña...")
    password_field = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='passwd']"))
    )
    password_field.send_keys(MOODLE_PASSWORD)
    
    next_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "idSIButton9"))
    )
    next_button.click()
    time.sleep(5)

    try:
        WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, "idBtn_Back"))
        ).click()
        print("✔ Saltado diálogo 'Mantener sesión'")
    except:
        print("✘ No apareció diálogo adicional")
    time.sleep(2)

    print("[4/6] Verificando acceso a Moodle...")
    WebDriverWait(driver, 20).until(
        EC.url_contains("moodle.isur.edu.pe")
    )
    print("✔ Login exitoso")
    time.sleep(3)

    print("[5/6] Accediendo a Página Principal...")
    try:
        pagina_principal = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 
                'a.nav-link.active[role="menuitem"][href="https://moodle.isur.edu.pe/"]'
            ))
        )
        pagina_principal.click()
        print("✔ Página Principal accedida")
        time.sleep(3)
    except Exception as e:
        print(f"✘ Error al acceder a Página Principal: {str(e)}")
        driver.save_screenshot("error_pagina_principal.png")
        raise

    print("[6/6] Accediendo a Office365...")
    try:
        office365_btn = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 
                'a.btn.login-identityprovider-btn[href*="auth/oidc/?source=loginpage"]'
            ))
        )
        office365_btn.click()
        print("✔ Botón Office365 clickeado")
        
        WebDriverWait(driver, 15).until(
            EC.url_contains("office.com") | EC.url_contains("microsoftonline.com")
        )
        print("✔ Redirección a Office365 confirmada")
        time.sleep(5)
        
    except Exception as e:
        print(f"✘ Error al acceder a Office365: {str(e)}")
        driver.save_screenshot("error_office365.png")
        raise

    print("\n>>> Accediendo a Mis Cursos <<<")
    try:
        driver.get("https://moodle.isur.edu.pe/")
        print("✔ Regresando a Moodle ISUR")
        time.sleep(3)
        
        mis_cursos_btn = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 
                'a.nav-link.active[href="https://moodle.isur.edu.pe/my/courses.php"]'
            ))
        )
        mis_cursos_btn.click()
        print("✔ Botón 'Mis Cursos' clickeado")
        
        WebDriverWait(driver, 15).until(
            EC.url_contains("moodle.isur.edu.pe/my/courses.php")
        )
        print("✔ Página 'Mis Cursos' cargada correctamente")
        
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".coursebox"))
        )
        print("✔ Cursos detectados en la página")
        time.sleep(5)
        
    except Exception as e:
        print(f"✘ Error al acceder a Mis Cursos: {str(e)}")
        driver.save_screenshot("error_mis_cursos.png")
        raise

    print("\n=== AUTOMATIZACIÓN COMPLETADA ===")
    print("Esperando 30 segundos antes de cerrar...")
    time.sleep(30)

except Exception as e:
    print(f"\nERROR: {str(e)}")
    driver.save_screenshot("error_final.png")
    time.sleep(30)

finally:
    driver.quit()
    print("\n=== NAVEGADOR CERRADO ===")