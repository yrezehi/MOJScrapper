from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import pandas as pd
from selenium.webdriver.common.action_chains import ActionChains
import os

# Setup WebDriver (replace with the path to your WebDriver)
driver = webdriver.Chrome()

print("[LOADING URL]")
driver.get("https://app.powerbi.com/view?r=eyJrIjoiNGI5OWM4NzctMDExNS00ZTBhLWIxMmYtNzIyMTJmYTM4MzNjIiwidCI6IjMwN2E1MzQyLWU1ZjgtNDZiNS1hMTBlLTBmYzVhMGIzZTRjYSIsImMiOjl9")
print("[LOADING URL FINISH]")

print("[INITIAL LOAD WAIT]")
time.sleep(8)
print("[INITIAL LOAD FINISHED]")

# Open table tab of the data in the Power BI report
driver.find_element(By.CSS_SELECTOR, "#pvExplorationHost > div > div > exploration > div > explore-canvas > div > div.canvasFlexBox > div > div.displayArea.disableAnimations.fitToPage > div.visualContainerHost.visualContainerOutOfFocus > visual-container-repeat > visual-container-group:nth-child(2) > transform > div > div.vcGroupBody.themableBackgroundColor.themableBorderColorSolid > visual-container-group:nth-child(5) > transform > div > div.vcGroupBody.themableBackgroundColor.themableBorderColorSolid > visual-container-group:nth-child(2) > transform > div > div.vcGroupBody.themableBackgroundColor.themableBorderColorSolid > visual-container:nth-child(2) > transform > div > div.visualContent > div > div > visual-modern").click()

# slicer-restatement
drop_down_cities = driver.find_elements(By.CSS_SELECTOR, ".slicer-dropdown-menu")

if len(drop_down_cities) >= 3:
    drop_down_cities[3].click()
    time.sleep(3)
    drop_down_cities_rows = driver.find_elements(By.CSS_SELECTOR, ".slicerItemContainer")
    if len(drop_down_cities_rows) >= 4:
        drop_down_cities_rows[4].click()

print("[ROWS LOAD WAIT]")
time.sleep(5)
print("[ROWS LOAD FINISHED]")

# Locate the table container
#table_container = driver.find_element(By.CLASS_NAME, 'interactive-grid')

table_container = driver.find_element(By.CSS_SELECTOR, ".mid-viewport")

data_rows = []

print("[SCRAPE DATA START]")

total_fetched = 0

rows = table_container.find_elements(By.CSS_SELECTOR, ".mid-viewport div.row")  # Update this selector if necessary

for row in rows:
        time.sleep(1)  # Adjust based on your connection speed and loading time

        columns = row.find_elements(By.CSS_SELECTOR, ".tablixAlignCenter")
        row_data = [column.text for column in columns]
        print(f"[INSERT ROW DATA #{row_data[6]} OF TOTAL {total_fetched}]")

        total_fetched += 1

        df = pd.DataFrame([row_data])
        header = not os.path.exists('scraped_data.csv')
        df.to_csv('scraped_data.csv', mode='a', index=False, header=header)

driver.execute_script("arguments[0].scrollIntoView();", rows[-1])

time.sleep(8)  # Adjust based on your connection speed and loading time

# Scroll until no more new data
while True:
    # Extract data from current view
    rows = table_container.find_elements(By.CSS_SELECTOR, ".mid-viewport div.row")  # Update this selector if necessary

    # Extract data from rows, skip the header row and since the scoll stops at last row so we skip it on n+1 iteration
    for row in rows[1:]:

        columns = row.find_elements(By.CSS_SELECTOR, ".tablixAlignCenter")
        row_data = [column.text for column in columns]
        print(f"[INSERT ROW DATA #{row_data[6]} OF TOTAL {total_fetched} of {row_data[5]}]")

        total_fetched += 1

        df = pd.DataFrame([row_data])
        header = not os.path.exists('scraped_data.csv')
        df.to_csv('scraped_data.csv', mode='a', index=False, header=header)

    print(f"[SCROLLING STARTED]")

    # Scroll down to the bottom of the container
    driver.execute_script("arguments[0].scrollIntoView();", rows[-1])

    time.sleep(1)  # Adjust based on your connection speed and loading time

    print(f"[SCROLLING FINISHED]")

    # Check if new rows are loaded by comparing number of rows before and after scroll
    new_rows = table_container.find_elements(By.CSS_SELECTOR, "table_viewport_container")

    if len(new_rows) == len(rows):
        # If no new rows are loaded, break the loop
        break

# Close the WebDriver
driver.quit()

print("Data scraped and saved to scraped_data.csv")