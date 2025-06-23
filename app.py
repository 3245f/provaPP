from flask import Flask, request, render_template, send_file, abort, redirect, url_for
import pandas as pd
import os
from datetime import datetime
# import threading # se si usa excel_lock per il file principale
import logging   
import requests  # Importa la libreria requests per le chiamate HTTP (necessaria per upload esterni)
import urllib.parse # Per codificare gli URL 

app = Flask(__name__)



# Configura il logging: imposta il livello minimo dei messaggi da visualizzare (INFO e superiori) e il formato del messaggio (timestamp, livello, messaggio)
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
EXCEL_FILE = "skills_trial.xlsx"


# Configura la modalità di invio delle risposte (E-MAIL/SHAREPOINT)
# In fase iniziale è impostata unicamente la modalità che consente l'invio tramite email
# Se si imposta su 'False' --> modalità SharePoint
# Controllo tramite variabile d'ambiente su Render
EMAIL_ONLY_MODE = os.environ.get("EMAIL_ONLY_MODE", "True").lower() == "true"

# Dettagli per la modalità solo email
DESTINATARIO_EMAIL = os.environ.get("DESTINATARIO_EMAIL", "stefania.giordano@alten.it") # da sostituire
#print(f"VERIFICA: DESTINATARIO_EMAIL utilizzato: '{DESTINATARIO_EMAIL}'") 
OGGETTO_EMAIL = os.environ.get("OGGETTO_EMAIL", "Modulo Competenze Alten")
#RICEVENTE_EMAIL = "stefania.giordano@alten.it"
#OGGETTO_EMAIL = "Modulo Competenze Alten"

# Configurazione SharePoint (todo)
GENERIC_SHAREPOINT_API_KEY = os.environ.get("GENERIC_SHAREPOINT_API_KEY", "YOUR_API_KEY")
GENERIC_SHAREPOINT_UPLOAD_API_URL = os.environ.get("GENERIC_SHAREPOINT_UPLOAD_API_URL", "https://your.genericsite.com/api/upload") 
# URL della cartella SharePoint che gli utenti dovrebbero vedere nel browser (todo)
SHAREPOINT_FOLDER_BROWSER_URL = os.environ.get("SHAREPOINT_FOLDER_BROWSER_URL", "https://your.sharepoint.com/sites/YourSite/SharedDocuments/YourFolder")


# Funzione per caricare un file sullo SharePoint (todo)
def upload_file_to_generic_sharepoint(file_path, file_name):
    logging.info(f"Tentativo di upload di '{file_name}' sullo SharePoint")
    headers = {
        "Authorization": f"Bearer {GENERIC_SHAREPOINT_API_KEY}", 
    }
    try:
        with open(file_path, 'rb') as f: 
            response = requests.put(GENERIC_SHAREPOINT_UPLOAD_API_URL + f"/{file_name}", headers=headers, data=f)
            response.raise_for_status() 
        logging.info(f"File '{file_name}' caricato su SharePoint generico con successo. Risposta: {response.status_code}")
        return True
    except FileNotFoundError:
        logging.error(f"Errore: File '{file_path}' non trovato per l'upload a SharePoint generico.")
        return False
    except requests.exceptions.RequestException as e:
        logging.error(f"Errore durante l'upload del file a SharePoint generico: {e}. Risposta: {getattr(e.response, 'text', 'Nessuna risposta testuale')}")
        return False



# Crea la directory USER_FILES_DIR se non esiste già dove exist_ok=True evita un errore se la directory esiste già
USER_FILES_DIR = "skills_user"
os.makedirs(USER_FILES_DIR, exist_ok=True)



user_df = pd.DataFrame(columns=[
         "Nome", "Email"])

if not os.path.exists(EXCEL_FILE):


    df = pd.DataFrame(columns=[
        "ID", "Nome", "Email"])

    df.to_excel(EXCEL_FILE, index=False)


# Funzione per assegnare un nuovo ID a ciascun nuovo utente che compila il Form
# L'ID sarà solo sequenziale per i nomi dei file temporanei, non persistente nel file principale
_current_id = 0
def get_next_id():
    global _current_id
    _current_id += 1
    return _current_id


# Per la versione con file excel generale
def remove_user_from_main_file(user_id):
    # Rimuovi la riga dal file principale basata sull'ID dell'utente
    if os.path.exists(EXCEL_FILE):
        df = pd.read_excel(EXCEL_FILE)
        # Trova e rimuovi la riga corrispondente all'ID dell'utente
        df = df[df["ID"] != user_id]
        df.to_excel(EXCEL_FILE, index=False)


# Funzione per aggiungere i dettagli di una specifica area al dizionario dei dati
def aggiungi_sezione(nome_sezione, scelte, dettagli_dict,data):
    data[f"Aree progetti {nome_sezione}"] = ", ".join(scelte)  
    # Aggiunge la colonna con i dettagli subito dopo la relativa sezione
    for area in dettagli_dict:
        data[area] = "\n\n".join(dettagli_dict[area]) if dettagli_dict[area] else ""
       


# Definisce la rotta principale dell'applicazione ("/")
# methods=["GET", "POST"] indica che la rotta gestisce sia le richieste GET (per visualizzare la pagina)
# che le richieste POST (per inviare i dati del form)
@app.route("/", methods=["GET", "POST"])

# Gestisce l'invio del form e il salvataggio dei dati
def index():
     # Inizializza a None ce successivo setaggio con valori veri in caso di invio POST riuscito
    success_message = None
    show_delete_button = False
    user_id = None  # Variabile per salvare l'ID dell'utente
    user_filename = None
    nome_utente = "" # Inizializza per essere sicuro che sia sempre definito

    # Controlla se la richiesta HTTP è di tipo POST (cioè, il form è stato inviato)      
    if request.method == "POST":

        # Genera il prossimo ID disponibile per il nuovo utente
        user_id = get_next_id()
        # Preleva i dati dal form dove "" indica che viene fornito un valore di default vuoto se il campo non è presente
        nome = request.form.get("nome", "")
        nome_utente = nome # Salva il nome per usarlo nel filename
        email = request.form.get("email", "")
        istruzione = request.form.get("istruzione", "")
        studi = request.form.get("studi", "")
        certificati = request.form.get("certificati", "")
        sede = request.form.get("sede", "")
        esperienza = request.form.get("esperienza", "")
        esperienza_alten = request.form.get("esperienza_alten", "")
        clienti_railway= request.form.getlist("clienti")  
        clienti_str = ", ".join(clienti_railway) if clienti_railway else "" 
        area_railway= request.form.getlist("area_railway")  
        area_str = ", ".join(area_railway) if area_railway else "" 
        normative = request.form.get("normative", "")
        metodologia= request.form.getlist("metodologia")  
        metodologia_str = ", ".join(metodologia) if metodologia else "" 
        sistemi_operativi = request.form.get("SistemiOperativi", "")
        altro= request.form.getlist("altro")  
        altro_str = ", ".join(altro) if altro else "" 
        hobby= request.form.getlist("hobby")  
        hobby_str = ", ".join(hobby) if hobby else "" 


        # Elaborazione delle sezioni dinamiche dei "Progetti" 
        # Ogni blocco segue una logica simile:
        # 1. Recupera le scelte generali per la categoria di progetto (es. 'sviluppo').
        # 2. Inizializza un dizionario `dettagli_<categoria>` con liste vuote per ogni sotto-area.
        # 3. Itera su ogni sotto-area:
        #    a. Se la sotto-area è stata selezionata nel form:
        #    b. Recupera tutti i campi specifici per quell'area (es. linguaggi, tool, durata, descrizione).
        #    c. Crea una lista di stringhe `esperienze`, dove ogni stringa è una combinazione dei dettagli per una singola esperienza.
        #    d. Assegna questa lista di esperienze al dizionario `dettagli_<categoria>[area]`.

# Progetti SVILUPPO
        progetti_sviluppo_si_no = request.form.get('progetti_sw_hw_auto', 'No')  
        scelte_progetti_sviluppo = request.form.getlist('sviluppo')  
        dettagli_sviluppo = {area: "" for area in ["Applicativi", "Firmware", "Web", "Mobile", "Scada", "Plc"]}
        for area in dettagli_sviluppo.keys():
            if area not in scelte_progetti_sviluppo:  
                continue  

            linguaggi = request.form.getlist(f'linguaggi_{area.lower()}[]')
            tool = request.form.getlist(f'tool_{area.lower()}[]')
            ambito = request.form.getlist(f'Ambito_{area.lower()}[]')
            nome_azienda = request.form.getlist(f'nome_azienda_{area.lower()}[]') # MODIFICA
            durata = request.form.getlist(f'durata_{area.lower()}[]')
            descrizione = request.form.getlist(f'descrizione_{area.lower()}[]')
            #print(f"{area} -> Linguaggi: {linguaggi}, Tool: {tool}, Ambito: {ambito}, Durata: {durata}, Descrizione: {descrizione}")
            esperienze = []
            for i in range(max(len(linguaggi), len(tool), len(ambito), len(durata), len(descrizione),len(nome_azienda)) ): # MODIFICA
                l = linguaggi[i] if i < len(linguaggi) else ""
                t = tool[i] if i < len(tool) else ""
                a = ambito[i] if i < len(ambito) else ""
                na = nome_azienda[i] if i < len(nome_azienda) else "" # MODIFICA: recupera il nome dell'azienda
                e = durata[i] if i < len(durata) else ""
                d = descrizione[i] if i < len(descrizione) else ""
                # Se l'ambito è "Aziendale", include il nome dell'azienda tra parentesi ()
                if a == "Aziendale" and na: # MODIFICA: Condizione per includere il nome dell'azienda
                    a = f"{a} ({na})"

                esperienze.append(f"{l} | {t} | {a} | {e} | {d}")
            dettagli_sviluppo[area] =esperienze

      
# Progetti V&V
        scelte_progetti_vv = request.form.getlist('v&v')  
        dettagli_vv = {area: "" for area in ["functional_testing", "test_and_commisioning", "unit", "analisi_statica", "analisi_dinamica", "automatic_test", "piani_schematici", "procedure", "cablaggi", "FAT", "SAT", "doc"]}
        for area in dettagli_vv.keys():
            if area not in scelte_progetti_vv:  
                continue  

            tecnologie = request.form.getlist(f'tecnologie_{area}[]')
            ambito = request.form.getlist(f'azienda_{area}[]')
            durata = request.form.getlist(f'durata_{area}[]')
            descrizione = request.form.getlist(f'descrizione_{area}[]')
            #print(f"{area} -> Tecnologie: {tecnologie}, Ambito: {ambito}, Durata: {durata}, Descrizione: {descrizione}")
            esperienze = []
            for i in range(max(len(tecnologie), len(ambito), len(descrizione))):
                t = tecnologie[i] if i < len(tecnologie) else ""
                a = ambito[i] if i < len(ambito) else ""
                e = durata[i] if i < len(durata) else ""
                d = descrizione[i] if i < len(descrizione) else ""
                esperienze.append(f" {t} | {a} | {e} | {d}")
            dettagli_vv[area] =esperienze
       
# Progetti System
        scelte_progetti_system = request.form.getlist('system')  
        dettagli_system = {area: "" for area in ["requirement_management", "requirement_engineering", "system_engineering", "project_engineering"]}
        for area in dettagli_system.keys():
            if area not in scelte_progetti_system:  
                continue  
            tecnologie = request.form.getlist(f'tecnologie_{area}[]')
            ambito = request.form.getlist(f'azienda_{area}[]')
            durata = request.form.getlist(f'durata_{area}[]')
            descrizione = request.form.getlist(f'descrizione_{area}[]')
            #print(f"{area} -> Tecnologie: {tecnologie}, Ambito: {ambito}, Durata: {durata}, Descrizione: {descrizione}")
            esperienze = []
            for i in range(max(len(tecnologie), len(ambito), len(descrizione))):
                t = tecnologie[i] if i < len(tecnologie) else ""
                a = ambito[i] if i < len(ambito) else ""
                e = durata[i] if i < len(durata) else ""
                d = descrizione[i] if i < len(descrizione) else ""
                esperienze.append(f"{t} | {a} | {e} | {d}")
            dettagli_system[area] =esperienze
        #print(dettagli_system)
       

# Progetti Safety
        scelte_progetti_safety = request.form.getlist('safety')  
        dettagli_safety = {area: "" for area in ["RAMS", "hazard_analysis", "verification_report", "fire_safety", "reg_402"]}
        #print(request.form) 
        for area in dettagli_safety.keys():
            if area not in scelte_progetti_safety: 
                continue  

            tecnologie = request.form.getlist(f'tecnologie_{area}[]')
            ambito = request.form.getlist(f'azienda_{area}[]')
            durata = request.form.getlist(f'durata_{area}[]')
            descrizione = request.form.getlist(f'descrizione_{area}[]')
            #print(f"{area} -> Tecnologie: {tecnologie}, Ambito: {ambito}, Durata: {durata}, Descrizione: {descrizione}")
            esperienze = []
            for i in range(max(len(tecnologie), len(ambito), len(descrizione))):
                t = tecnologie[i] if i < len(tecnologie) else ""
                a = ambito[i] if i < len(ambito) else ""
                e = durata[i] if i < len(durata) else ""
                d = descrizione[i] if i < len(descrizione) else ""
                esperienze.append(f"{t} | {a} | {e} | {d}")
            dettagli_safety[area] =esperienze
        #print(dettagli_safety)
      

# Progetti Segnalamento      
        scelte_progetti_segnalamento = request.form.getlist('segnalamento')  
        dettagli_seg = {area: "" for area in ["piani_schematici_segnalamento", "cfg_impianti", "layout_apparecchiature", "architettura_rete", "computo_metrico"]}
        for area in dettagli_seg.keys():
            if area not in scelte_progetti_segnalamento:  
                continue  

            tecnologie = request.form.getlist(f'tecnologie_{area}[]')
            ambito = request.form.getlist(f'azienda_{area}[]')
            durata = request.form.getlist(f'durata_{area}[]')
            descrizione = request.form.getlist(f'descrizione_{area}[]')
            #print(f"{area} -> Tecnologie: {tecnologie}, Ambito: {ambito}, Durata: {durata}, Descrizione: {descrizione}")
            esperienze = []
            for i in range(max(len(tecnologie),len(ambito), len(descrizione))):
                t = tecnologie[i] if i < len(tecnologie) else ""
                a = ambito[i] if i < len(ambito) else ""
                e = durata[i] if i < len(durata) else ""
                d = descrizione[i] if i < len(descrizione) else ""
                esperienze.append(f"{t} | {a} | {e} | {d}")
            dettagli_seg[area] = esperienze
        #print(dettagli_seg)
       



# Progetti BIM
        progetti_bim_si_no = request.form.get('progetti_bim', 'No')  
        scelte_progetti_bim = request.form.getlist('bim')  
        #print(scelte_progetti_bim) 
        dettagli_bim = {area: "" for area in ["modellazione_e_digitalizzazione", "verifica_analisi_e_controllo_qualita", "gestione_coordinamento_e_simulazione", "visualizzazione_realtavirtuale_e_rendering"]}
        #print(request.form) 
        for area in dettagli_bim.keys():
            if area not in scelte_progetti_bim:  
                continue  
            tool = request.form.getlist(f'tool_{area}[]')
            azienda = request.form.getlist(f'azienda_{area}[]')
            durata = request.form.getlist(f'durata_{area}[]')
            descrizione = request.form.getlist(f'descrizione_{area}[]')
            certificazione = request.form.getlist(f'certificazioni_{area}[]')
            #print(f"{area} -> Tool: {tool}, Azienda: {azienda}, Durata: {durata}, Descrizione: {descrizione}, Certificazioni: {certificazione}")
            esperienze = []
            for i in range(max(len(certificazione), len(tool), len(azienda), len(descrizione))):
                t = tool[i] if i < len(tool) else ""
                a = azienda[i] if i < len(azienda) else ""
                e = durata[i] if i < len(durata) else ""
                d = descrizione[i] if i < len(descrizione) else ""
                c = certificazione[i] if i < len(certificazione) else ""
                esperienze.append(f" {t} | {a} | {e} | {d} | {c}")
            dettagli_bim[area] =esperienze
        #print("dettagli bim",dettagli_bim)
    


# Progetti PM
        progetti_pm_si_no = request.form.get('progetti_pm', 'No')  
        scelte_progetti_pm = request.form.getlist('pm')  
        dettagli_pm = {area: "" for area in ["project_manager_office", "project_manager", "risk_manager", "resource_manager", "quality_manager", "communication_manager", "portfolio_manager", "program_manager","team_leader", "business_analyst", "contract_back_office"]}
        #print(request.form) 
        for area in dettagli_pm.keys():
            if area not in scelte_progetti_pm:  
                continue  

            tool = request.form.getlist(f'tool_{area}[]')
            azienda = request.form.getlist(f'azienda_{area}[]')
            durata = request.form.getlist(f'durata_{area}[]')
            descrizione = request.form.getlist(f'descrizione_{area}[]')
            #print(f"{area} -> Tool: {tool}, Azienda: {azienda}, Durata: {durata}, Descrizione: {descrizione}")
            esperienze = []
            for i in range(max(len(tool), len(azienda), len(descrizione))):
                t = tool[i] if i < len(tool) else ""
                a = azienda[i] if i < len(azienda) else ""
                e = durata[i] if i < len(durata) else ""
                d = descrizione[i] if i < len(descrizione) else ""
                esperienze.append(f"{t} | {a} | {e} | {d}")
            dettagli_pm[area] =esperienze
        #print("dettagli pm",dettagli_pm)
       


# Crea un dizionario 'data' (che viene poi convertito in DataFrame pandas) con tutte le informazioni raccolte dal form che devono essere salvate nel file excel
        data = {
            "ID": user_id,
            "Nome": nome,
            "Email": email,
            "Istruzione": istruzione,
            "Indirizzo di studio": studi,
            "Sede Alten": sede,
            "Esperienza (anni)": esperienza,
            "Esperienza Alten (anni)": esperienza_alten,
            "Certificazioni": certificati,
            "Clienti Railway":  clienti_str, 
            "Area Railway": area_str, 
            "Normative": normative, 
            "Metodologie lavoro": metodologia_str,
            "Sistemi Operativi": sistemi_operativi,
            "Info aggiuntive": altro_str,
            "Hobby": hobby_str,
        }




        # Aggiunta delle varie sezioni di progetto con i dettagli in ordine al dizionario
        aggiungi_sezione("Sviluppo", scelte_progetti_sviluppo, dettagli_sviluppo,data)
        aggiungi_sezione("V&V", scelte_progetti_vv, dettagli_vv,data)
        aggiungi_sezione("Safety", scelte_progetti_safety, dettagli_safety,data)
        aggiungi_sezione("System", scelte_progetti_system, dettagli_system,data)
        aggiungi_sezione("Segnalamento", scelte_progetti_segnalamento, dettagli_seg,data)
        aggiungi_sezione("BIM", scelte_progetti_bim, dettagli_bim,data)
        aggiungi_sezione("Project Management", scelte_progetti_pm, dettagli_pm,data)   


    # Controlla se l'azione del form è "submit_main" (pulsante "Salva Risposta")
        if request.form['action'] == 'submit_main':
            try:
                # Salvataggio dei dati nel file Excel principale
                # with excel_lock:
                #     logging.info(f"Lock acquisito per la scrittura del file Excel principale.")
                #     df = pd.read_excel(EXCEL_FILE)
                #     for col in data.keys():
                #         if col not in df.columns:
                #             df[col] = ''
                #     df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)
                #     df.to_excel(EXCEL_FILE, index=False)
                #     logging.info(f"Dati scritti sul file Excel principale. {len(df)} righe totali. Lock rilasciato.")
                
                success_message = "Risposte inviate con successo!"

                # Generazione del nome del file con nome utente e data
                nome_unificato = "".join(c for c in nome_utente if c.isalnum() or c == '_').strip().replace(' ', '_') # Rimuovi spazi o caratteri speciali dal nome utente
                if not nome_unificato: # Se il nome è vuoto o solo caratteri speciali
                    nome_unificato = "Utente" # Nome di fallback
                
                user_filename = f"{nome_unificato}_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"
                logging.info(f"Nome del file utente generato: {user_filename}")
                user_filepath = os.path.join(USER_FILES_DIR, user_filename)
                logging.info(f"Percorso del file utente: {user_filepath}")
                user_df_single = pd.DataFrame([data])
                # ID escluso dal file scaricato dall'utente
                if 'ID' in user_df_single.columns:
                    user_df_single = user_df_single.drop(columns=["ID"])
                user_df_single.to_excel(user_filepath, index=False)
                
            except Exception as e:
                success_message = f'Si è verificato un errore durante l\'invio delle risposte: {e}'
                logging.error(f"Errore durante l'invio delle risposte o il salvataggio del file: {e}", exc_info=True)

        # Gestisce l'azione di esportazione al "SharePoint" (todo)
        elif request.form['action'] == 'export_to_generic_sharepoint':
            filename_to_export = request.form.get("user_filename_to_export")

            if filename_to_export:
                file_path_to_export = os.path.join(USER_FILES_DIR, filename_to_export)
                if os.path.exists(file_path_to_export):
                    logging.info(f"Tentativo di esportare '{filename_to_export}' a SharePoint generico.")
                    if upload_file_to_generic_sharepoint(file_path_to_export, filename_to_export):
                        success_message = f"File '{filename_to_export}' esportato su SharePoint generico con successo!"
                        logging.info(f"File '{filename_to_export}' caricato su SharePoint generico.")
                    else:
                        success_message = f"Errore nell'esportazione di '{filename_to_export}' a SharePoint generico."
                        logging.error(f"Fallito l'upload di '{filename_to_export}' a SharePoint generico.")
                else:
                    success_message = f"File '{filename_to_export}' non trovato per l'esportazione."
                    logging.warning(f"File '{filename_to_export}' non trovato per l'esportazione a SharePoint generico.")
            else:
                success_message = "Nome file per l'esportazione su SharePoint generico non specificato."
                logging.warning("Nessun nome file fornito per l'esportazione su SharePoint generico.")
            
            user_filename = filename_to_export     


    # Renderizza il template 'form.html', passando i messaggi di successo/errore,lo stato del pulsante di eliminazione e il nome del file utente per il download e le info sulle modalità di invio del file excel
    return render_template("form.html", 
        success_message=success_message, 
        show_delete_button=show_delete_button, 
        user_filename=user_filename,
        sharepoint_folder_browser_url=SHAREPOINT_FOLDER_BROWSER_URL,
        email_only_mode=EMAIL_ONLY_MODE, 
        destinatario_email=DESTINATARIO_EMAIL, 
        oggetto_email=OGGETTO_EMAIL) 


# Definisce la rotta per il download dei file personale
@app.route("/download")
def download():
    # Preleva il tipo di file da scaricare dal parametro 'file' nell'URL che di default è "main" (file principale)
    file_type = request.args.get("file", "main")  # Valore di default: 'main'
    # Se il tipo di file richiesto è "personal" 
    if file_type == "personal":
        filename = request.args.get("filename")  # Il nome del file personale
        # Se il nome del file non è stato fornito, genera un errore 400 (Bad Request)
        if not filename:
            return abort(400, description="Missing filename parameter")
        
        # Costruisce il percorso completo del file utente
        user_filepath = os.path.join(USER_FILES_DIR, filename)
        # Se il file utente non esiste, genera un errore 404 (Not Found)
        if not os.path.exists(user_filepath):
            return abort(404, description="File not found")

        # Invia il file utente come allegato per il download: as_attachment=True forza il browser a scaricare il file anziché visualizzarlo
        return send_file(user_filepath, as_attachment=True, download_name=filename)

    return abort(404, description="Invalid file type or file not found.") # Gestisce il caso di download di tipo "main" non più supportato
    


if __name__ == "__main__":
    app.run(debug=True)


if __name__ == "__main__":
    app.run(debug=True)
