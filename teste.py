def executar_visivelmente():
    print("[INFO] Iniciando navegador Chrome em modo VISÍVEL...")
    options = webdriver.ChromeOptions()

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    print("[DEBUG] ChromeDriver iniciado.")

    driver.get("https://google.com")  # teste básico
    print("[DEBUG] Página carregada.")

    input("[INFO] Pressione Enter para sair...")
    driver.quit()
