from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import pandas as pd

# Setup WebDriver (replace with the path to your WebDriver)
driver = webdriver.Chrome()

print("[LOADING URL]")
driver.get("https://app.powerbi.com/view?r=eyJrIjoiNGI5OWM4NzctMDExNS00ZTBhLWIxMmYtNzIyMTJmYTM4MzNjIiwidCI6IjMwN2E1MzQyLWU1ZjgtNDZiNS1hMTBlLTBmYzVhMGIzZTRjYSIsImMiOjl9")
print("[LOADING URL FINISH]")

print("[INITIAL LOAD WAIT]")
time.sleep(10)
print("[INITIAL LOAD FINISHED]")

# Open table tab of the data in the Power BI report
driver.find_element(By.CSS_SELECTOR, "#pvExplorationHost > div > div > exploration > div > explore-canvas > div > div.canvasFlexBox > div > div.displayArea.disableAnimations.fitToPage > div.visualContainerHost.visualContainerOutOfFocus > visual-container-repeat > visual-container-group:nth-child(2) > transform > div > div.vcGroupBody.themableBackgroundColor.themableBorderColorSolid > visual-container-group:nth-child(5) > transform > div > div.vcGroupBody.themableBackgroundColor.themableBorderColorSolid > visual-container-group:nth-child(2) > transform > div > div.vcGroupBody.themableBackgroundColor.themableBorderColorSolid > visual-container:nth-child(2) > transform > div > div.visualContent > div > div > visual-modern").click()

print("[ROWS LOAD WAIT]")
time.sleep(10)
print("[ROWS LOAD FINISHED]")

# Locate the table container
table_container = driver.find_element(By.CLASS_NAME, 'interactive-grid')

data_rows = []

print("[SCRAPE DATA START]")

while True:
    # Get current height of the container
    last_height = driver.execute_script("return arguments[0].scrollHeight", table_container)
    
    # Scroll to the bottom of the container
    driver.execute_script("arguments[0].scrollTo(0, arguments[0].scrollHeight);", table_container)

    print("[TABLE ROWS LOAD WAIT START]")
    # Wait for new data to load
    time.sleep(5)  # Adjust based on your connection speed and loading time
    print("[TABLE ROWS LOAD WAIT FINISHED]")

    # Get updated table rows
    rows = table_container.find_elements(By.CSS_SELECTOR, "div.innerContainer div.row")  # Update this selector if necessary

    # Extract data from rows
    for row in rows:
        columns = row.find_elements(By.CSS_SELECTOR, ".tablixAlignCenter")
        row_data = [column.text for column in columns]
        print("[INSERT ROW DATA]")
        data_rows.append(row_data)
    
    # Check if the height is the same as before scrolling
    new_height = driver.execute_script("return arguments[0].scrollHeight", table_container)
    if new_height == last_height:
        driver.execute_script("arguments[0].scrollTo(0, arguments[0].scrollHeight);", table_container)
        time.sleep(5)
        new_height = driver.execute_script("return arguments[0].scrollHeight", table_container)
        if new_height == last_height:
            break

# Close the WebDriver
driver.quit()

# Save data to a DataFrame and CSV
df = pd.DataFrame(data_rows)
df.to_csv('scraped_data.csv', index=False)

print("Data scraped and saved to scraped_data.csv")