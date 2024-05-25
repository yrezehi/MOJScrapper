from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import pandas as pd
from selenium.webdriver.common.action_chains import ActionChains

# Setup WebDriver (replace with the path to your WebDriver)
driver = webdriver.Chrome()

print("[LOADING URL]")
driver.get("https://app.powerbi.com/view?r=eyJrIjoiNGI5OWM4NzctMDExNS00ZTBhLWIxMmYtNzIyMTJmYTM4MzNjIiwidCI6IjMwN2E1MzQyLWU1ZjgtNDZiNS1hMTBlLTBmYzVhMGIzZTRjYSIsImMiOjl9")
print("[LOADING URL FINISH]")

print("[INITIAL LOAD WAIT]")
time.sleep(5)
print("[INITIAL LOAD FINISHED]")

# Open table tab of the data in the Power BI report
driver.find_element(By.CSS_SELECTOR, "#pvExplorationHost > div > div > exploration > div > explore-canvas > div > div.canvasFlexBox > div > div.displayArea.disableAnimations.fitToPage > div.visualContainerHost.visualContainerOutOfFocus > visual-container-repeat > visual-container-group:nth-child(2) > transform > div > div.vcGroupBody.themableBackgroundColor.themableBorderColorSolid > visual-container-group:nth-child(5) > transform > div > div.vcGroupBody.themableBackgroundColor.themableBorderColorSolid > visual-container-group:nth-child(2) > transform > div > div.vcGroupBody.themableBackgroundColor.themableBorderColorSolid > visual-container:nth-child(2) > transform > div > div.visualContent > div > div > visual-modern").click()

print("[ROWS LOAD WAIT]")
time.sleep(5)
print("[ROWS LOAD FINISHED]")

# Locate the table container
table_container = driver.find_element(By.CLASS_NAME, 'interactive-grid')

table_viewport_container = driver.find_element(By.CSS_SELECTOR, ".mid-viewport")
last_height = driver.execute_script("return arguments[0].scrollHeight", table_viewport_container)

data_rows = []

print("[SCRAPE DATA START]")

# Scroll until no more new data
while True:
    # Extract data from current view
    rows = table_container.find_elements(By.CSS_SELECTOR, "div.innerContainer div.row")  # Update this selector if necessary

    # Extract data from rows
    for row in rows:
        columns = row.find_elements(By.CSS_SELECTOR, ".tablixAlignCenter")
        row_data = [column.text for column in columns]
        if row_data not in data_rows:  # Only add new rows to avoid duplicates
            print(f"[INSERT ROW DATA #{row_data[6]}]")
            data_rows.append(row_data)

    # Scroll down to the bottom of the container
    driver.execute_script(f"arguments[0].scrollTo(0, {last_height});", table_viewport_container)

    # Wait to load the page
    time.sleep(10)  # Adjust based on your connection speed and loading time

    # Check if new rows are loaded by comparing number of rows before and after scroll
    new_rows = table_container.find_elements(By.CSS_SELECTOR, "table_viewport_container")
    if len(new_rows) == len(rows):
        # If no new rows are loaded, break the loop
        break
    last_height = driver.execute_script("return arguments[0].scrollHeight", table_viewport_container)

# Close the WebDriver
    
driver.quit()

# Save data to a DataFrame and CSV
df = pd.DataFrame(data_rows)
df.to_csv('scraped_data.csv', index=False)

print("Data scraped and saved to scraped_data.csv")