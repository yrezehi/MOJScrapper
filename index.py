from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import pandas as pd

# Setup WebDriver (replace with the path to your WebDriver)
driver_path = '/path/to/your/webdriver'
driver = webdriver.Chrome(driver_path)

# Open the Power BI report
url = 'https://app.powerbi.com/view?r=eyJrIjoiNGI5OWM4NzctMDExNS00ZTBhLWIxMmYtNzIyMTJmYTM4MzNjIiwidCI6IjMwN2E1MzQyLWU1ZjgtNDZiNS1hMTBlLTBmYzVhMGIzZTRjYSIsImMiOjl9'
driver.get(url)

# Allow the page to load
time.sleep(10)  # Adjust based on your connection speed

# Locate the table container
table_container = driver.find_element(By.CLASS_NAME, 'interactive-grid')

data_rows = []

while True:
    # Get current height of the container
    last_height = driver.execute_script("return arguments[0].scrollHeight", table_container)
    
    # Scroll to the bottom of the container
    driver.execute_script("arguments[0].scrollTo(0, arguments[0].scrollHeight);", table_container)
    
    # Wait for new data to load
    time.sleep(5)  # Adjust based on your connection speed and loading time
    
    # Get updated table rows
    rows = table_container.find_elements(By.CSS_SELECTOR, "div.innerContainer div.row")  # Update this selector if necessary
    
    # Extract data from rows
    for row in rows:
        columns = row.find_elements(By.CLASS_NAME, "cellContent")
        row_data = [column.text for column in columns]
        data_rows.append(row_data)
    
    # Check if the height is the same as before scrolling
    new_height = driver.execute_script("return arguments[0].scrollHeight", table_container)
    if new_height == last_height:
        break

# Close the WebDriver
driver.quit()

# Save data to a DataFrame and CSV
df = pd.DataFrame(data_rows)
df.to_csv('scraped_data.csv', index=False)

print("Data scraped and saved to scraped_data.csv")