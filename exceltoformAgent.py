import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.webdriver import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

def fill_google_form(form_url, excel_file):
    # Load responses from Excel file
    df = pd.read_excel(excel_file)  # Read data
    total_submissions = len(df)  # Number of rows (submissions)
    
    # Setup Chrome WebDriver
    options = Options()
    options.add_argument("--start-maximized")  # Maximize window
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    for i, row in df.iterrows():
        driver.get(form_url)
        time.sleep(3)  # Allow page to load

        try:
            # Get all radio button questions
            all_questions = driver.find_elements(By.XPATH, "//div[@role='radiogroup']")
            
            for idx, question in enumerate(all_questions):
                radio_options = question.find_elements(By.XPATH, ".//div[@role='radio']")

                if radio_options:
                    answer = str(row.iloc[idx]).strip()  # Get answer from Excel
                    selected_option = None

                    # Match the answer with available options
                    for option in radio_options:
                        if option.get_attribute("aria-label").strip().lower() == answer.lower():
                            selected_option = option
                            break
                    
                    if selected_option:
                        # Scroll into view & Click using JavaScript
                        driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", selected_option)
                        time.sleep(0.5)  # Small delay for stability
                        driver.execute_script("arguments[0].click();", selected_option)

            # Submit the form
            time.sleep(1)
            submit_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Submit')]"))
            )
            submit_button.click()

            print(f"Form submitted {i+1}/{total_submissions} times.")
            time.sleep(3)  # Wait before restarting

        except Exception as e:
            print(f"Error encountered: {e}")

    driver.quit()

# Example Usage
form_link = "https://forms.gle/2ZTSqHEec2M6Tzdy5"
excel_file = "testing_ai.xlsx"  # Ensure this file exists
fill_google_form(form_link, excel_file)
