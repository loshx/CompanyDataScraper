import pandas as pd
import os
import tkinter as tk
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import customtkinter as ctk
import re
import time
import random
from selenium.webdriver.common.keys import Keys

# Setări optimizate pentru browser
options = Options()
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--enable-fast-unload")
options.add_experimental_option("prefs", {"profile.managed_default_content_settings.images": 2})

COLUMNS = [
    "term_search", "Company Name", "IDNO", "Status", "Registration date",
    "Physical adress", "Legal address", "Nr. of employees",
    "Contact 1", "Contact 2", "Contact 3", "Administrator", "Activity",
    "Net Sales (2024)", "Net Profit (2024)", "Employees (2024)",
    "Net Sales (2023)", "Net Profit (2023)", "Employees (2023)",
    "Net Sales (2022)", "Net Profit (2022)", "Employees (2022)"
]

def save_to_excel_with_formatting(df, file_path):
    """Salvează DataFrame-ul în Excel cu formatare îmbunătățită"""
    try:
        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Company Data', index=False)
            
            workbook = writer.book
            worksheet = writer.sheets['Company Data']
            
            # Formate personalizate
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#4F81BD',
                'font_color': 'white',
                'border': 1,
                'font_size': 12
            })
            
            cell_format = workbook.add_format({
                'text_wrap': True,
                'valign': 'top',
                'border': 1,
                'font_size': 10
            })
            
            # Aplică formatarea header-elor
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Aplică formatarea celulelor
            for row_num in range(1, len(df) + 1):
                for col_num in range(len(df.columns)):
                    worksheet.write(row_num, col_num, df.iloc[row_num-1, col_num], cell_format)
            
            # Ajustează lățimea coloanelor
            for i, col in enumerate(df.columns):
                max_length = max(
                    df[col].astype(str).apply(len).max(),
                    len(col)
                ) + 3
                worksheet.set_column(i, i, min(max_length, 50))
            
            # Adaugă filtre
            worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)
            
            # Îngheță prima linie (header)
            worksheet.freeze_panes(1, 0)
            
    except Exception as e:
        print(f"Eroare la salvarea fișierului Excel: {e}")
        # Fallback la metoda simplă
        df.to_excel(file_path, index=False)

def search_companies_via_google():
    search_queries = entry.get("1.0", tk.END).strip().split("\n")
    search_queries = [q.strip() for q in search_queries if q.strip()]

    if not search_queries:
        messagebox.showerror("Eroare", "Introduceți cel puțin un nume de companie!")
        return

    # Configurare browser optimizată
    options = Options()
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
    options.add_argument("--lang=en-US,en;q=0.9")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 3)

    file_path = "company_data.xlsx"
    first_time = not os.path.exists(file_path)
    scraped_data = {}
    all_results = []

    # Încărcăm datele existente
    if not first_time:
        df_existing = pd.read_excel(file_path)
        for _, row in df_existing.iterrows():
            term = str(row["term_search"]).strip()
            if term not in scraped_data:
                scraped_data[term] = row.to_dict()

    for query in search_queries:
        if query in scraped_data:
            print(f"ℹ️ Folosim date existente pentru: {query}")
            all_results.append(scraped_data[query])
            continue

        company_data = {col: "N/A" for col in COLUMNS}
        company_data["term_search"] = query

        try:
            # Căutare Google cu specificația pentru versiunea engleză
            driver.get("https://www.google.com/search?hl=en")
            time.sleep(random.uniform(1, 2))

            search_box = wait.until(EC.presence_of_element_located((By.NAME, "q")))
            search_box.clear()
            search_box.send_keys(f"{query} data2b")
            search_box.send_keys(Keys.RETURN)
            time.sleep(random.uniform(2, 3))

            # Găsire link către data2b.md/en/
            try:
                first_result = wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//a[contains(@href,'data2b.md/ro/search?')]")
                ))
                profile_url = first_result.get_attribute("href")
                print(f"Link găsit (RO): {profile_url}")
                
                # Convertim link-ul RO în EN dacă este posibil
                profile_url = profile_url.replace('/ro/', '/en/')
                print(f"Încercăm versiunea EN: {profile_url}")
                
            except Exception as e:
                print(f"Nu s-a găsit link pentru: {query} - {str(e)}")
                all_results.append(company_data)
                continue

            # Accesare profil companie
            driver.get(profile_url)
            time.sleep(random.uniform(2, 3))

            # Verificare dacă există un link mai specific
            try:
                more_specific_link = wait.until(EC.presence_of_element_located(
                    (By.XPATH, "//a[contains(@href, '/companies/')]")
                ))
                if more_specific_link:
                    specific_url = more_specific_link.get_attribute("href")
                    
                    # Forțăm versiunea EN dacă link-ul specific are /ro/
                    if '/ro/' in specific_url:
                        specific_url = specific_url.replace('/ro/', '/en/')
                        print(f"Convertit link specific la EN: {specific_url}")
                        
                    if specific_url != profile_url:
                        print(f"Găsit link mai specific: {specific_url}")
                        driver.get(specific_url)
                        time.sleep(random.uniform(1, 2))
            except:
                print("Nu s-a găsit un link mai specific, continuăm cu link-ul inițial")

            # Verificare finală și corecție URL
            current_url = driver.current_url
            if '/ro/' in current_url:
                print(f"Corectăm URL din RO în EN: {current_url}")
                current_url = current_url.replace('/ro/', '/en/')
                driver.get(current_url)
                time.sleep(random.uniform(1, 2))
            elif '/en/' not in current_url:
                print(f"Redirecționare neașteptată la: {current_url}")
                all_results.append(company_data)
                continue

            # EXTRAGERE NUME COMPANIE - MODIFICAT (h1 în loc de h2)
            try:
                company_name = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "h1.MuiTypography-h2"))).text.strip()
                company_data["Company Name"] = company_name
            except:
                try:
                    company_name = wait.until(EC.presence_of_element_located((By.TAG_NAME, "h1"))).text.strip()
                    company_data["Company Name"] = company_name
                except:
                    try:
                        company_name = wait.until(EC.presence_of_element_located((By.TAG_NAME, "h2"))).text.strip()
                        company_data["Company Name"] = company_name
                    except:
                        company_data["Company Name"] = "N/A"

            # Extragere date din tabel (aceeași logică ca în search_companies)
            rows = driver.find_elements(By.CSS_SELECTOR, ".MuiTableRow-root")
            for row in rows:
                try:
                    key = row.find_element(By.TAG_NAME, "th").text.strip()
                    value = row.find_element(By.TAG_NAME, "td").text.strip()
                    if key in COLUMNS:
                        company_data[key] = value
                except:
                    continue

            # Adresă fizică
            try:
                addr_elem = WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.XPATH, "//th[contains(text(), 'Physical address')]/following-sibling::td"))
                )
                company_data["Physical adress"] = addr_elem.text.strip() or "N/A"
            except:
                pass

            # Contacte
            phone_numbers = driver.find_elements(By.CSS_SELECTOR, 'a[href^="tel:"]')[:3]
            for idx, phone in enumerate(phone_numbers):
                company_data[f"Contact {idx+1}"] = phone.text.strip() or "N/A"

            # Administrator
            try:
                director_elements = driver.find_elements(By.CSS_SELECTOR, ".MuiListItemText-root span")
                for director in director_elements:
                    name = director.text.strip()
                    if "[Administrator]" in name:
                        company_data["Administrator"] = name.replace(" [Administrator]", "")
                        break
            except:
                pass

            # Activitate
            try:
                activity = driver.find_element(By.XPATH, "//div[.//span[text()='Main activity']]//span[contains(@class, 'MuiTypography-body2')]").text.strip()
                company_data["Activity"] = activity
            except:
                try:
                    activity = driver.find_element(By.XPATH, "//div[.//span[text()='Activities']]//span[contains(@class, 'MuiTypography-body2')]").text.strip()
                    company_data["Activity"] = activity
                except:
                    pass

            # Performanță financiară
            performance_rows = driver.find_elements(By.CSS_SELECTOR, "tbody.MuiTableBody-root tr")
            available_years = {}
            for row in performance_rows:
                try:
                    cells = row.find_elements(By.TAG_NAME, "td")
                    if len(cells) == 4:
                        year = cells[0].text.strip()
                        if year in ["2024", "2023", "2022"]:
                            available_years[year] = {
                                "Net Sales": cells[1].text.strip(),
                                "Net Profit": cells[2].text.strip(),
                                "Employees": cells[3].text.strip()
                            }
                except:
                    continue

            for y in ["2024", "2023", "2022"]:
                if y in available_years:
                    company_data[f"Net Sales ({y})"] = available_years[y]["Net Sales"]
                    company_data[f"Net Profit ({y})"] = available_years[y]["Net Profit"]
                    company_data[f"Employees ({y})"] = available_years[y]["Employees"]

            scraped_data[query] = company_data
            all_results.append(company_data)
            print(f"✅ Date salvate pentru: {query}")

        except Exception as e:
            print(f"❌ Eroare la procesarea {query}: {str(e)}")
            all_results.append(company_data)

    driver.quit()

    if all_results:
        df = pd.DataFrame(all_results)
        # Salvăm în același fișier ca search_companies()
        if os.path.exists(file_path):
            df_existing = pd.read_excel(file_path)
            df = pd.concat([df_existing, df]).drop_duplicates(subset=["term_search"], keep="last")
        
        # SALVARE CU FORMATARE ÎMBUNĂTĂȚITĂ
        save_to_excel_with_formatting(df, file_path)
        
        new_entries = len([q for q in search_queries if q not in scraped_data])
        messagebox.showinfo(
            "Succes", 
            f"Fișier actualizat cu succes!\n"
            f"Total intrări: {len(all_results)}\n"
            f"Date noi adăugate: {new_entries}\n"
            f"Duplicate păstrate: {len(search_queries) - new_entries}"
        )
    else:
        messagebox.showinfo("Eroare", "Nu s-au putut obține date.")

def search_companies():
    search_queries = entry.get("1.0", tk.END).strip().split("\n")
    search_queries = [q.strip() for q in search_queries if q.strip()]

    if not search_queries:
        messagebox.showerror("Eroare", "Introduceți cel puțin un nume de companie!")
        return

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 5)

    file_path = "company_data.xlsx"
    first_time = not os.path.exists(file_path)

    scraped_data = {}
    all_results = []

    if not first_time:
        df_existing = pd.read_excel(file_path)
        for _, row in df_existing.iterrows():
            term = str(row["term_search"]).strip()
            if term not in scraped_data:
                scraped_data[term] = row.to_dict()

    for query in search_queries:
        if query in scraped_data:
            print(f"ℹ️ Folosim date existente pentru: {query}")
            all_results.append(scraped_data[query])
            continue

        url = f"https://www.data2b.md/en/search?q={query}"
        driver.get(url)

        company_data = {col: "N/A" for col in COLUMNS}
        company_data["term_search"] = query

        try:
            retry_count = 1
            while retry_count > 0:
                try:
                    first_result = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[data-testid="result-0"]')))
                    first_result.click()

                    if "#google_vignette" in driver.current_url:
                        print(f"⚠️ Reclama detectată pentru {query}. Reîncercare...")
                        driver.refresh()
                        retry_count -= 1
                    else:
                        break
                except:
                    driver.refresh()
                    retry_count -= 1

            if retry_count == 0:
                print(f"❌ Nu s-a putut accesa pagina pentru: {query}")
                all_results.append(company_data)
                continue

            # EXTRAGERE NUME COMPANIE - MODIFICAT (h1 în loc de h2)
            try:
                company_name = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "h1.MuiTypography-h2"))).text.strip()
                company_data["Company Name"] = company_name
            except:
                try:
                    company_name = wait.until(EC.presence_of_element_located((By.TAG_NAME, "h1"))).text.strip()
                    company_data["Company Name"] = company_name
                except:
                    company_data["Company Name"] = "N/A"

            rows = driver.find_elements(By.CSS_SELECTOR, ".MuiTableRow-root")
            for row in rows:
                try:
                    key = row.find_element(By.TAG_NAME, "th").text.strip()
                    value = row.find_element(By.TAG_NAME, "td").text.strip()
                    if key in COLUMNS:
                        company_data[key] = value
                except:
                    continue

            try:
                addr_elem = WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.XPATH, "//th[contains(text(), 'Physical address')]/following-sibling::td"))
                )
                company_data["Physical adress"] = addr_elem.text.strip() or "N/A"
            except:
                pass

            phone_numbers = driver.find_elements(By.CSS_SELECTOR, 'a[href^="tel:"]')[:3]
            for idx, phone in enumerate(phone_numbers):
                company_data[f"Contact {idx+1}"] = phone.text.strip() or "N/A"

            try:
                director_elements = driver.find_elements(By.CSS_SELECTOR, ".MuiListItemText-root span")
                for director in director_elements:
                    name = director.text.strip()
                    if "[Administrator]" in name:
                        company_data["Administrator"] = name.replace(" [Administrator]", "")
                        break
            except:
                pass

            try:
                activity = driver.find_element(By.XPATH, "//div[.//span[text()='Main activity']]//span[contains(@class, 'MuiTypography-body2')]").text.strip()
                company_data["Activity"] = activity
            except:
                try:
                    activity = driver.find_element(By.XPATH, "//div[.//span[text()='Activities']]//span[contains(@class, 'MuiTypography-body2')]").text.strip()
                    company_data["Activity"] = activity
                except:
                    pass

            performance_rows = driver.find_elements(By.CSS_SELECTOR, "tbody.MuiTableBody-root tr")
            available_years = {}
            for row in performance_rows:
                try:
                    cells = row.find_elements(By.TAG_NAME, "td")
                    if len(cells) == 4:
                        year = cells[0].text.strip()
                        if year in ["2024", "2023", "2022"]:
                            available_years[year] = {
                                "Net Sales": cells[1].text.strip(),
                                "Net Profit": cells[2].text.strip(),
                                "Employees": cells[3].text.strip()
                            }
                except:
                    continue

            for y in ["2024", "2023", "2022"]:
                if y in available_years:
                    company_data[f"Net Sales ({y})"] = available_years[y]["Net Sales"]
                    company_data[f"Net Profit ({y})"] = available_years[y]["Net Profit"]
                    company_data[f"Employees ({y})"] = available_years[y]["Employees"]

            scraped_data[query] = company_data
            all_results.append(company_data)
            print(f"✅ Date salvate pentru: {query}")

        except Exception as e:
            print(f"❌ Eroare la procesarea: {query} - {str(e)}")
            all_results.append(company_data)

    driver.quit()

    if all_results:
        df = pd.DataFrame(all_results)
        # SALVARE CU FORMATARE ÎMBUNĂTĂȚITĂ
        save_to_excel_with_formatting(df, file_path)

        new_entries = len([q for q in search_queries if q not in scraped_data])
        messagebox.showinfo(
            "Succes", 
            f"Fișier actualizat cu succes!\n"
            f"Total intrări: {len(all_results)}\n"
            f"Date noi adăugate: {new_entries}\n"
        )
    else:
        messagebox.showinfo("Eroare", "Nu s-au putut obține date.")

#SECOND FUNCTION
def extract_company_data(url, page, driver):
    full_url = f"{url}&page={page}"
    driver.get(full_url)

    from time import sleep
    sleep(2)

    ids = []

    # Căutăm toate cardurile de companie
    cards = driver.find_elements(By.XPATH, "//div[contains(@class, 'MuiPaper-root')]")
    print(f"🔍 Găsite {len(cards)} carduri pe pagina {page}")

    for idx, card in enumerate(cards):
        try:
            # Verificăm dacă există un SVG în card
            try:
                svg = card.find_element(By.TAG_NAME, "svg")
                # dacă găsește svg – continuă
            except:
                continue  # nu are SVG → sărim

            # Caută IDNO în orice paragraf care conține doar cifre
            idno_text = None
            for p in card.find_elements(By.TAG_NAME, "p"):
                txt = p.text.strip()
                if re.match(r'^[1-9]\d*$', txt):
                    idno_text = txt
                    break

            if idno_text:
                ids.append(idno_text)
                print(f"[{idx+1}] ✅ IDNO extras: {idno_text}")
            else:
                print(f"[{idx+1}] ⚠️ Card cu SVG dar fără IDNO valid")
        except Exception as e:
            print(f"[{idx+1}] ❌ Eroare la extragere: {e}")

    return ids

def start_extraction():
    """ Funcție pentru a extrage IDNO-urile companiilor din paginile specificate """
    
    url = url_entry.get().strip()
    num_pages = pages_entry.get().strip()

    if not url:
        messagebox.showerror("Eroare", "Introduceți un URL valid!")
        return

    if not num_pages.isdigit():
        messagebox.showerror("Eroare", "Introduceți un număr valid de pagini!")
        return

    num_pages = int(num_pages)
    progress_label.configure(text="🔄 Se extrag date...", text_color="yellow")
    root.update_idletasks()

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    all_ids = []
    for page in range(1, num_pages + 1):
        try:
            all_ids.extend(extract_company_data(url, page, driver))  
            progress_label.configure(text=f"📄 Pagina {page} procesată...", text_color="blue")
            root.update_idletasks()
        except Exception as e:
            print(f"⚠️ Eroare la pagina {page}: {str(e)}")

    driver.quit()

    df = pd.DataFrame(all_ids, columns=["Company ID"])
    # SALVARE CU FORMATARE ÎMBUNĂTĂȚITĂ
    save_to_excel_with_formatting(df, "company_ids.xlsx")

    progress_label.configure(text="✅ Extracție finalizată!", text_color="green")
    messagebox.showinfo("Succes", "Extracția IDNO-urilor s-a terminat. Datele au fost salvate în 'company_ids.xlsx'!")

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

root = ctk.CTk()
root.title("Căutare Companii și Extragere Date")
root.geometry("600x400")

def show_page(page):
    if page == "page1":
        page1_frame.pack(fill="both", expand=True)
        page2_frame.pack_forget()
    elif page == "page2":
        page2_frame.pack(fill="both", expand=True)
        page1_frame.pack_forget()

page1_frame = ctk.CTkFrame(root)

title_label = ctk.CTkLabel(page1_frame, text="🔍 Căutare Companii", font=("Arial", 22, "bold"))
title_label.pack(pady=10)

input_frame = ctk.CTkFrame(page1_frame)
input_frame.pack(pady=5, padx=20, fill="both", expand=True)

entry = ctk.CTkTextbox(input_frame, font=("Arial", 14), height=120, wrap="word")
entry.pack(padx=10, pady=5, fill="both", expand=True)

search_button = ctk.CTkButton(page1_frame, text="🔍 Caută și Salvează", font=("Arial", 14, "bold"),
                              command=search_companies, corner_radius=10, fg_color="blue")
search_button.pack(pady=15)

nav_to_page2_button = ctk.CTkButton(page1_frame, text="Go to Company IDNO by URL", font=("Arial", 12),
                                    command=lambda: show_page("page2"), corner_radius=10, fg_color="gray")
nav_to_page2_button.pack(pady=10)

google_search_button = ctk.CTkButton(page1_frame, text="🔍 Caută via Google", font=("Arial", 14, "bold"),
                                     command=search_companies_via_google, corner_radius=10, fg_color="green")
google_search_button.pack(pady=5)

page2_frame = ctk.CTkFrame(root)

company_title_label = ctk.CTkLabel(page2_frame, text="🔗 Company IDNO by URL", font=("Arial", 22, "bold"))
company_title_label.pack(pady=20)

url_frame = ctk.CTkFrame(page2_frame)
url_frame.pack(fill='x', padx=20, pady=5)

url_label = ctk.CTkLabel(url_frame, text="Enter URL:", font=("Arial", 14))
url_label.pack(side='left', padx=(0, 10))
url_entry = ctk.CTkEntry(url_frame, width=300)
url_entry.pack(side='left', expand=True, fill='x')

pages_frame = ctk.CTkFrame(page2_frame)
pages_frame.pack(fill='x', padx=20, pady=5)

pages_label = ctk.CTkLabel(pages_frame, text="Enter number of pages:", font=("Arial", 14))
pages_label.pack(side='left', padx=(0, 10))
pages_entry = ctk.CTkEntry(pages_frame, width=100)
pages_entry.pack(side='left')

extract_button = ctk.CTkButton(page2_frame, text="🚀 Start Extraction", command=start_extraction,
                               font=("Arial", 14, "bold"), corner_radius=10, fg_color="blue")
extract_button.pack(pady=15)

nav_to_page1_button = ctk.CTkButton(page2_frame, text="Go to Căutare Companii", font=("Arial", 12),
                                    command=lambda: show_page("page1"), corner_radius=10, fg_color="gray")
nav_to_page1_button.pack(pady=10)

progress_label = ctk.CTkLabel(root, text="", font=("Arial", 14))
progress_label.pack(pady=10)

show_page("page1")
root.mainloop()