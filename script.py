import os
from docx import Document
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

def connexion(driver):
    url_login = ""
    driver.get(url_login)
    username = ""  
    password = ""
    try:
        champ_username = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "login")) 
        )
        champ_password = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "password"))
        )
        bouton_connexion = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "bg-dendreo-primary")) 
        )
        champ_username.clear()
        champ_password.clear()
        champ_username.send_keys(username) 
        champ_password.send_keys(password)
        bouton_connexion.click()
    except Exception as e:
        print(f"Erreur de connexion : {e}")

def lister_fichiers(dossier):
    fichiers = []
    for fichier in os.listdir(dossier):
        if os.path.isfile(os.path.join(dossier, fichier)):
            nom_sans_underscore = fichier.replace("_", " ")  
            nom_sans_extension = nom_sans_underscore.rsplit(".", 1)[0]
            fichiers.append(nom_sans_extension)  
    return fichiers

def extraire_titres_existants(driver):
    titres = set()  # Utiliser un ensemble pour éviter les doublons
    catalogue = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "catalogue"))
    )
    catalogue.click()

    while True:
        # Attendre que tous les éléments soient présents
        WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, "dd-mini-link--module"))
        )
        time.sleep(1)  # Ajouter une courte pause pour être sûr que tous les éléments se chargent

        # Collecter les titres uniques
        elements = driver.find_elements(By.CLASS_NAME, "dd-mini-link--module")
        for element in elements:
            titre = element.get_attribute("title")
            if titre:  # Ajouter le titre seulement s'il n'est pas vide
                titres.add(titre)

        # Vérifier si le bouton "Suivant" est activé
        try:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(1)  # Pause pour permettre au bouton de devenir cliquable
            bouton_suivant = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "dataTableBuilder_next"))
            )

            # Vérifier si le bouton est désactivé, puis arrêter la boucle
            classe_bouton = bouton_suivant.get_attribute("class")
            if classe_bouton and "disabled" in classe_bouton:
                print("Fin de la pagination. Le bouton 'Suivant' est désactivé.")
                break
            else:
                bouton_suivant.click()
                WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.CLASS_NAME, "dd-mini-link--module"))
                )
        except Exception as e:
            print("Erreur lors de la pagination :", e)
            break

    titres_liste = list(titres)  # Convertir l'ensemble en liste pour le retour
    print("Titres existants trouvés :", titres_liste)
    return titres_liste

def lire_contenu_docx(chemin_fichier):
    document = Document(chemin_fichier)
    contenu = []
    for para in document.paragraphs:
        contenu.append(para.text)
    return '\n'.join(contenu)

def ajout_intitule(driver, fichier):
    driver.get("")
    intitule = driver.find_element(By.ID, "intitule")
    intitule.clear()
    intitule.send_keys(fichier)

def ajout_description(driver, fichier, dossier):
    chemin_fichier = os.path.join(dossier, fichier) + ".docx"
    chemin_fichier = chemin_fichier.replace(" ", "_")
    contenu = lire_contenu_docx(chemin_fichier)
    driver.switch_to.frame(driver.find_element(By.ID, "description_ifr"))
    champ_contenu = driver.find_element(By.TAG_NAME, "body")
    champ_contenu.clear()
    champ_contenu.send_keys(contenu)
    driver.switch_to.default_content()

def valider(driver):
    time.sleep(2)
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(2)
    Valider = driver.find_element(By.ID, "sub_add_module")
    Valider.click()

def mettre_en_forme(driver):
    bouton_assistant_ai = WebDriverWait(driver, 60).until(
        EC.element_to_be_clickable((By.XPATH, "//div[@id='p_description']//button[@data-mce-name='dd_ai_assistant']"))
    )
    bouton_assistant_ai.click()
    option_mettre_en_forme = WebDriverWait(driver, 60).until(
        EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Mettre en forme')]"))
    )
    option_mettre_en_forme.click()
    bouton_remplacer = WebDriverWait(driver, 60).until(
        EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Remplacer']"))
    )
    bouton_remplacer.click()

def activer_ia(driver, div_id):
    bouton_assistant_ai = WebDriverWait(driver, 60).until(
        EC.element_to_be_clickable((By.XPATH, f"//div[@id='{div_id}']//button[@data-mce-name='dd_ai_assistant']"))
    )
    bouton_assistant_ai.click()
    option_generer_contenu = WebDriverWait(driver, 60).until(
        EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Générer du contenu')]"))
    )
    option_generer_contenu.click()
    bouton_remplacer = WebDriverWait(driver, 60).until(
        EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='Remplacer']"))
    )
    bouton_remplacer.click()
    time.sleep(2)

if __name__ == "__main__":
    dossier = "./Files/"
    driver = webdriver.Firefox()

    try:
        connexion(driver)
        fichiers_a_traiter = lister_fichiers(dossier)
        titres_existants = extraire_titres_existants(driver)
        fichiers_a_ajouter = [f for f in fichiers_a_traiter if f not in titres_existants]
        
        for fichier in fichiers_a_ajouter:
            ajout_intitule(driver, fichier)
            ajout_description(driver, fichier, dossier)
            mettre_en_forme(driver)
            time.sleep(3)
            activer_ia(driver, "p_")
            activer_ia(driver, "p_")
            activer_ia(driver, "p_")
            activer_ia(driver, "p_")
            activer_ia(driver, "p_")
            # input("continuer?")
            valider(driver)
    finally:
        driver.quit()
