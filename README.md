# SSACaste
# Find Caste of the Students Using SSA Number Need to Excel in Download foloder and the excel name must be  RptStudentList.xlsx
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import UnexpectedAlertPresentException, NoAlertPresentException
from webdriver_manager.chrome import ChromeDriverManager
import time

# 1. Path to your file
file_path = r"C:\Users\SAI\Downloads\RptStudentList.xlsx"
df = pd.read_excel(file_path)

if 'Social Category' not in df.columns:
    df['Social Category'] = ""

# 2. Setup Chrome
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.get("https://www.ssgujarat.org/CTELogin.aspx")

wait = WebDriverWait(driver, 10)
original_window = driver.current_window_handle

# --- MASSIVE SAFETY NET STARTS HERE ---
try:
    # 3. Automation Loop
    for index, row in df.iterrows():
        
        # Skip empty rows
        raw_uid = row['Student UID']
        if pd.isna(raw_uid) or str(raw_uid).strip() == "":
            continue 
            
        uid_str = str(raw_uid).split('.')[0].strip()
        
        # --- NEW FEATURE: AUTO-RESUME ---
        # If this student already has a valid category saved, skip them!
        existing_data = str(row['Social Category']).strip()
        if existing_data and existing_data not in ["nan", "Not Found", "Alert Blocked", "Invalid UID", ""]:
            print(f"Skipping {uid_str}: Already completed ({existing_data})")
            continue
        
        try:
            search_box = driver.find_element(By.ID, "TxtSearch")
            search_box.clear()
            search_box.send_keys(uid_str)
            
            driver.find_element(By.ID, "BtnSearch").click()
            
            # Check for invalid UID pop-ups
            try:
                alert = driver.switch_to.alert
                alert_text = alert.text
                alert.accept() 
                print(f"Skipped UID {uid_str}: Website rejected it ({alert_text})")
                df.at[index, 'Social Category'] = "Invalid UID"
                
                # --- NEW FEATURE: AUTOSAVE ---
                df.to_excel(file_path, index=False)
                continue 
            except NoAlertPresentException:
                pass 

            # Wait for new tab
            wait.until(EC.number_of_windows_to_be(2))
            
            for window_handle in driver.window_handles:
                if window_handle != original_window:
                    driver.switch_to.window(window_handle)
                    break
            
            # Grab category
            category_element = wait.until(EC.presence_of_element_located((By.ID, "lblsocialcategory")))
            caste_text = category_element.text
            if caste_text == "":
                caste_text = category_element.get_attribute("innerText")
                
            df.at[index, 'Social Category'] = caste_text
            print(f"Success: {uid_str} -> '{caste_text}'")
            
            driver.close()
            driver.switch_to.window(original_window)
            
        except UnexpectedAlertPresentException:
            print(f"Failed UID {uid_str}: Pop-up blocked the script.")
            df.at[index, 'Social Category'] = "Alert Blocked"
            try:
                driver.switch_to.alert.accept()
            except:
                pass
                
        except Exception as e:
            error_type = type(e).__name__
            print(f"Failed UID {uid_str}. Error: {error_type}")
            df.at[index, 'Social Category'] = "Not Found"
            
            # Cleanup tabs
            try:
                driver.switch_to.alert.accept()
            except:
                pass
            while len(driver.window_handles) > 1:
                driver.switch_to.window(driver.window_handles[-1])
                driver.close()
            driver.switch_to.window(original_window)
            driver.get("https://www.ssgujarat.org/CTELogin.aspx")

        # --- NEW FEATURE: AUTOSAVE ---
        # Saves the file silently after every single student is checked
        df.to_excel(file_path, index=False)

# Catch manual stops (Ctrl+C) or massive unhandled crashes
except KeyboardInterrupt:
    print("\n--- You manually paused/stopped the script. ---")
except Exception as fatal_error:
    print(f"\n--- A critical error stopped the script: {fatal_error} ---")

# --- EMERGENCY SAVE PROTOCOL ---
# This 'finally' block runs no matter what happens above!
finally:
    print("\nExecuting final save and closing Chrome...")
    df.to_excel(file_path, index=False)
    driver.quit()
    print("All scraped data is safely secured in your Excel file!")
