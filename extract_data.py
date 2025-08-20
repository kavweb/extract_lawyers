from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import time
import random

input_file = "links.xlsx"  
output_file = "lawyers_data.xlsx"  

df = pd.read_excel(input_file)
links = df.iloc[:, 0].tolist()  

#  Selenium
options = webdriver.ChromeOptions()
options.add_argument("--headless")  
driver = webdriver.Chrome(options=options)


data = []


for index, link in enumerate(links):
    try:
        
        driver.get(link)
        time.sleep(random.uniform(2, 5)) 
        
        name = driver.find_element(By.CLASS_NAME, "text-lg.font-bold").text.strip()

        try:
            rank = driver.find_element(By.XPATH, "//strong/following-sibling::span").text.strip()
        except:
            rank = "نامشخص"
            print(driver.page_source)  

        license_number = driver.find_element(By.XPATH, "//span[contains(text(),'شماره پروانه:')]/following-sibling::span").text.strip()
        validity = driver.find_element(By.XPATH, "//span[contains(text(),'اعتبار:')]/following-sibling::span").text.strip()
        province = driver.find_element(By.XPATH, "//span[contains(text(),'استان:')]/following-sibling::span").text.strip()
        city = driver.find_element(By.XPATH, "//span[contains(text(),'شهر:')]/following-sibling::span").text.strip()
        phone = driver.find_element(By.XPATH, "//span[contains(text(),'تلفن همراه:')]/following-sibling::span").text.strip()

        try:
            office_phone = driver.find_element(By.XPATH, "//span[contains(text(),'تلفن موسسه:')]/following-sibling::span").text.strip()
            if office_phone == "":
                office_phone = "نامشخص"
        except:
            office_phone = "نامشخص"

        address = driver.find_element(By.XPATH, "//span[contains(text(),'آدرس موسسه:')]/following-sibling::span").text.strip()

        data.append([name, rank, license_number, validity, province, city, phone, office_phone, address])
        print(f"{index+1}/{len(links)} - {name} استخراج شد ✅")
    
    except Exception as e:
        print(f"{index+1}/{len(links)} - خطا در پردازش لینک {link}: {e}")

df_output = pd.DataFrame(data, columns=["نام", "رتبه", "شماره پروانه", "اعتبار", "استان", "شهر", "تلفن همراه", "تلفن موسسه", "آدرس"])
df_output.to_excel(output_file, index=False)

print("✅ اطلاعات با موفقیت ذخیره شد.")
driver.quit()

