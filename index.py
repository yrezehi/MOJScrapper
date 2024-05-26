from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import pandas as pd
from selenium.webdriver.common.action_chains import ActionChains
import os


DEFAULT_LOAD_WAIT_TIME = 8
SHORT_LOAD_WAIT_TIME = 3
VERY_SHORT_LOAD_WAIT_TIME = 1

driver = webdriver.Chrome()

print("[LOADING URL]")
driver.get("https://app.powerbi.com/view?r=eyJrIjoiNGI5OWM4NzctMDExNS00ZTBhLWIxMmYtNzIyMTJmYTM4MzNjIiwidCI6IjMwN2E1MzQyLWU1ZjgtNDZiNS1hMTBlLTBmYzVhMGIzZTRjYSIsImMiOjl9")

print("[INITIAL LOAD WAIT]")
time.sleep(DEFAULT_LOAD_WAIT_TIME)

# Open table tab of the data
driver.find_element(By.CSS_SELECTOR, "#pvExplorationHost > div > div > exploration > div > explore-canvas > div > div.canvasFlexBox > div > div.displayArea.disableAnimations.fitToPage > div.visualContainerHost.visualContainerOutOfFocus > visual-container-repeat > visual-container-group:nth-child(2) > transform > div > div.vcGroupBody.themableBackgroundColor.themableBorderColorSolid > visual-container-group:nth-child(5) > transform > div > div.vcGroupBody.themableBackgroundColor.themableBorderColorSolid > visual-container-group:nth-child(2) > transform > div > div.vcGroupBody.themableBackgroundColor.themableBorderColorSolid > visual-container:nth-child(2) > transform > div > div.visualContent > div > div > visual-modern").click()

CITY_DROPDOWN_ELEMENT_INDEX = 3
CITY_DROPDOWN_ELEMENT_ITEM_INDEX = 4

# Pick city from the dropdown and the type of property
drop_down_cities = driver.find_elements(By.CSS_SELECTOR, ".slicer-dropdown-menu")
if len(drop_down_cities) >= CITY_DROPDOWN_ELEMENT_INDEX:
    drop_down_cities[CITY_DROPDOWN_ELEMENT_INDEX].click()
    time.sleep(SHORT_LOAD_WAIT_TIME)
    drop_down_cities_rows = driver.find_elements(By.CSS_SELECTOR, ".slicerItemContainer")
    if len(drop_down_cities_rows) >= CITY_DROPDOWN_ELEMENT_ITEM_INDEX:
        drop_down_cities_rows[CITY_DROPDOWN_ELEMENT_ITEM_INDEX].click()

YEAR_DROPDOWN_ELEMENT_INDEX = 3
YEAR_DROPDOWN_ELEMENT_ITEM_INDEXS = [1, 2]

# Pick city from the dropdown and the type of property
drop_down_years = driver.find_elements(By.CSS_SELECTOR, ".slicer-dropdown-menu")
if len(drop_down_years) >= YEAR_DROPDOWN_ELEMENT_INDEX:
    drop_down_years[YEAR_DROPDOWN_ELEMENT_INDEX].click()
    time.sleep(SHORT_LOAD_WAIT_TIME)
    drop_down_years = driver.find_elements(By.CSS_SELECTOR, ".slicerItemContainer")
    for year_index in YEAR_DROPDOWN_ELEMENT_ITEM_INDEXS:
        if len(drop_down_years) >= year_index:
            drop_down_years[year_index].click()

print("[TABLE'S ROWS LOAD WAIT]")
time.sleep(SHORT_LOAD_WAIT_TIME)

# Element of the table
table_container = driver.find_element(By.CSS_SELECTOR, ".mid-viewport")
data_rows = []
total_fetched = 0

print("[SCRAPE DATA START]")

rows = table_container.find_elements(By.CSS_SELECTOR, ".mid-viewport div.row")  # Update this selector if necessary

for row in rows:
        time.sleep(VERY_SHORT_LOAD_WAIT_TIME) 

        columns = row.find_elements(By.CSS_SELECTOR, ".tablixAlignCenter")
        row_data = [column.text for column in columns]
        print(f"[INSERT ROW DATA #{row_data[6]} OF TOTAL {total_fetched}]")

        total_fetched += 1

        df = pd.DataFrame([row_data])
        header = not os.path.exists('scraped_data.csv')
        df.to_csv('scraped_data.csv', mode='a', index=False, header=header)

driver.execute_script("arguments[0].scrollIntoView();", rows[-1])

time.sleep(DEFAULT_LOAD_WAIT_TIME)  

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

    time.sleep(VERY_SHORT_LOAD_WAIT_TIME)  # Adjust based on your connection speed and loading time

    # Check if new rows are loaded by comparing number of rows before and after scroll
    new_rows = table_container.find_elements(By.CSS_SELECTOR, "table_viewport_container")

    if len(new_rows) == 0:
        break

# Close the WebDriver
driver.quit()

print("Data scraped and saved to scraped_data.csv")