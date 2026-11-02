import os
import time
from datetime import datetime
import win32com.client as win32
from docx import Document
from selenium import webdriver

def telecomanda_word_e_salva(percorso_file):
    print("[*] Avvio Microsoft Word in background per modificare metadati...")
    try:
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0 
        
        doc = word.Documents.Open(percorso_file)
        
        print("[*] Documento aperto. Simulazione di un uso normale mantenendolo aperto per 200 secondi...")
        time.sleep(200) 
        
        doc.Content.InsertAfter(" ")
        
        doc.Save()
        doc.Close()
        word.Quit()

        tempo_finale = time.time()
        os.utime(percorso_file, (tempo_finale, tempo_finale))

        print("[*] Metadati e statistiche aggiornati con successo da MS Word.")
    except Exception as e:
        print(f"[!] Errore con pywin32: {e}")

def genera_alibi():
    print("[*] Apertura browser e navigazione su Wikipedia in corso...")
    
    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True) 
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_argument("--disable-blink-features=AutomationControlled")
    
    driver = webdriver.Chrome(options=options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    
    try:
        driver.get("https://it.wikipedia.org/wiki/Catania")
        print("[*] Pagina caricata. Simulazione di lettura (attesa di 15 secondi)...")
        time.sleep(15) 
    except Exception as e:
        print(f"[!] Errore durante la navigazione web: {e}")
    finally:
        driver.quit()
        print("[*] Finestra del browser chiusa.")

    print("[*] Generazione del documento Word in corso...")
    
    doc = Document()
    utente_windows = os.getlogin() 
    adesso = datetime.utcnow() 
    
    prop = doc.core_properties
    prop.author = utente_windows
    prop.last_modified_by = utente_windows
    prop.comments = "" 
    prop.created = adesso 
    prop.modified = adesso 

    doc.add_heading('Weekend a Catania', 0)
    doc.add_paragraph("Itinerario per giorno 8 novembre:")
    doc.add_paragraph("Piazza del Duomo e l'Elefante (Liotru)", style='List Number')
    doc.add_paragraph("Il mercato della Pescheria", style='List Number')
    doc.add_paragraph("Via Etnea per lo shopping e la vista sull'Etna", style='List Number')
    doc.add_paragraph("Monastero dei Benedettini", style='List Number')
    
    doc.add_paragraph(
        "Ricordarsi di provare qualche arancino "
        "al ragù e la granita con brioche."
    )
    
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    file_path = os.path.join(desktop_path, "Itinerario_Catania.docx")
    
    doc.save(file_path)
    print(f"[*] Salvataggio iniziale su: {file_path}")

    telecomanda_word_e_salva(file_path)
    print("[*] Generazione alibi completata.")

if __name__ == "__main__":
    genera_alibi()