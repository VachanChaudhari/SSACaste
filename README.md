# SSACaste
# Find Caste of the Students Using SSA Number Need to Excel in Download foloder and the excel name must be  RptStudentList.xlsx
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time

# 1. Path to your file
file_path = r"C:\Users\SAI\Downloads\RptStudentList.xlsx"
df = pd.read_excel(file_path)

# Create the column if it doesn't exist
if 'Social Category' not in df.columns:
    df['Social Category'] = ""

# 2. Setup Chrome
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.get("https://www.ssgujarat.org/CTELogin.aspx")

wait = WebDriverWait(driver, 10)
original_window = driver.current_window_handle # Saves the main search tab

# 3. Automation Loop - PROCESSING ALL ROWS
for index, row in df.iterrows():
    try:
        # Find the search box and type the UID
        search_box = driver.find_element(By.ID, "TxtSearch")
        search_box.clear()
        search_box.send_keys(str(row['Student UID']))
        
        # Click Go
        driver.find_element(By.ID, "BtnSearch").click()
        
        # Wait for the new tab to open
        wait.until(EC.number_of_windows_to_be(2))
        
        # Switch Selenium's focus to the new tab
        for window_handle in driver.window_handles:
            if window_handle != original_window:
                driver.switch_to.window(window_handle)
                break
        
        # Grab the caste data
        category_element = wait.until(EC.presence_of_element_located((By.ID, "lblsocialcategory")))
        
        caste_text = category_element.text
        if caste_text == "":
            caste_text = category_element.get_attribute("innerText")
            
        # Save to the dataframe and print success
        df.at[index, 'Social Category'] = caste_text
        print(f"Success: {row['Student UID']} -> '{caste_text}'")
        
        # Close the result tab and switch back to the main search tab
        driver.close()
        driver.switch_to.window(original_window)
        
    except Exception as e:
        error_type = type(e).__name__
        print(f"Failed UID {row['Student UID']}. Error: {error_type}")
        df.at[index, 'Social Category'] = "Not Found"
        
        # Cleanup: If it crashes, close extra tabs and reset to main tab
        while len(driver.window_handles) > 1:
            driver.switch_to.window(driver.window_handles[-1])
            driver.close()
        driver.switch_to.window(original_window)
        driver.get("https://www.ssgujarat.org/CTELogin.aspx")

# 4. Save and Close
print("Saving data to Excel file... Please wait.")
df.to_excel(file_path, index=False)
driver.quit()
print("Process perfectly finished! All students have been processed.")
