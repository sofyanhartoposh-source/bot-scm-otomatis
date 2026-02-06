import time
import requests
import json
import pandas as pd
import io
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def jalankan_bot():
    # 1. Konfigurasi Browser Siluman
    options = uc.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    
    driver = uc.Chrome(options=options)
    
    try:
        print("Membuka SCM PLN Nusa Daya...")
        driver.get("https://scm.nusadaya.net/login")
        
        wait = WebDriverWait(driver, 20)
        
        # 2. Proses Login
        print("Memasukkan kredensial...")
        email_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='text' or @placeholder='Email atau NIP']")))
        pass_input = driver.find_element(By.XPATH, "//input[@type='password']")
        
        email_input.send_keys("sofyan.hartopo@pln.co.id")
        pass_input.send_keys("9314014DY")
        
        login_btn = driver.find_element(By.XPATH, "//button[contains(text(), 'Log in')]")
        login_btn.click()
        
        print("Login berhasil, menunggu session...")
        time.sleep(10) # Kasih waktu buat server bikin session cookie
        
        # 3. Ambil Cookie Login
        # Ini penting supaya kita bisa download file lewat link export
        session_cookies = {c['name']: c['value'] for c in driver.get_cookies()}
        
        # 4. Download Excel lewat link Export Langsung
        export_url = "https://scm.nusadaya.net/monitoring-kontrak-rinci/export?khs=all&bidang=all&tahun=2026&stage="
        print("Mendownload data via Export Link...")
        
        r = requests.get(export_url, cookies=session_cookies)
        
        # 5. Baca Excel pake Pandas (Tanpa simpan file ke laptop)
        print("Membaca isi Excel...")
        df = pd.read_excel(io.BytesIO(r.content))
        
        # Bersihkan data (ganti NaN jadi string kosong biar gak error di JSON)
        df = df.fillna("")
        
        # Ambil SEMUA baris data
        all_data = df.values.tolist()
        
        # 6. Kirim Ke Google Sheets (Ditimpa Total)
        gas_url = "https://script.google.com/macros/s/AKfycbzuWeZNwdAqMDcg81bfkiOhIosljdswcR0tdRWeenQjwokhgyuM_0PZORvMt8IH5-E/exec" 
        payload = json.dumps({"rows": all_data})
        
        headers = {'Content-Type': 'application/json'}
        print(f"Mengirim {len(all_data)} baris data ke Google Sheets...")
        
        response = requests.post(gas_url, data=payload, headers=headers)
        print(f"Respon Google Sheets: {response.text}")

    except Exception as e:
        print(f"Waduh, ada kendala: {str(e)}")
    finally:
        driver.quit()
        print("Bot selesai bekerja.")

if __name__ == "__main__":
    jalankan_bot()
