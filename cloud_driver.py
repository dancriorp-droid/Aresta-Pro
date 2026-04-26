"""
cloud_driver.py — Driver helper para Streamlit Community Cloud
Centraliza a criação do Selenium WebDriver usando Chromium headless.
Substitui todas as referências a Edge/msedgedriver do código original.
"""

import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options


def criar_driver(headless=True):
    """
    Cria e retorna um WebDriver Chrome/Chromium configurado para rodar
    no Streamlit Community Cloud (Linux headless).
    
    Funciona tanto no Streamlit Cloud quanto localmente se o Chrome estiver instalado.
    """
    options = Options()
    
    # === Flags obrigatórias para ambiente cloud/headless ===
    if headless:
        options.add_argument("--headless=new")
    
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--blink-settings=imagesEnabled=false")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-software-rasterizer")
    options.add_argument("--disable-background-timer-throttling")
    options.add_argument("--disable-backgrounding-occluded-windows")
    options.add_argument("--disable-renderer-backgrounding")
    
    # User-agent para evitar bloqueio
    options.add_argument(
        "user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
    
    # === Detecta caminho do Chromium/ChromeDriver ===
    # Streamlit Cloud (Debian/Ubuntu via packages.txt)
    chrome_paths = [
        "/usr/bin/chromium-browser",
        "/usr/bin/chromium",
        "/usr/bin/google-chrome",
        "/usr/bin/google-chrome-stable",
    ]
    driver_paths = [
        "/usr/bin/chromedriver",
        "/usr/lib/chromium-browser/chromedriver",
        "/usr/lib/chromium/chromedriver",
    ]
    
    chrome_bin = None
    for p in chrome_paths:
        if os.path.exists(p):
            chrome_bin = p
            break
    
    driver_bin = None
    for p in driver_paths:
        if os.path.exists(p):
            driver_bin = p
            break
    
    if chrome_bin:
        options.binary_location = chrome_bin
    
    if driver_bin:
        service = Service(executable_path=driver_bin)
    else:
        # Fallback: deixa o Selenium encontrar automaticamente
        service = Service()
    
    driver = webdriver.Chrome(service=service, options=options)
    driver.set_page_load_timeout(60)
    driver.implicitly_wait(5)
    
    return driver
