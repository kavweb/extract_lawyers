import time
from openpyxl import Workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service


def extract_all_pages(url, out_xlsx="links.xlsx"):
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(options=options)

    all_links = []

    try:
        driver.get(url)
        time.sleep(2)
        input('Select filter then press Enter:')
        page_num = 1
        while True:
            print(f"Extracting page: {page_num} ...")
            time.sleep(2)

            last_height = driver.execute_script("return document.body.scrollHeight")
            while True:
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(1)
                new_height = driver.execute_script("return document.body.scrollHeight")
                if new_height == last_height:
                    break
                last_height = new_height

            div = driver.find_element(By.CSS_SELECTOR, "div.flex.flex-col.grid-cols-3.gap-8.md\\:grid")
            links = div.find_elements(By.TAG_NAME, "a")
            hrefs = [a.get_attribute("href") for a in links if a.get_attribute("href")]

            all_links.extend(hrefs)
            print(f"✅ {len(hrefs)} find.")

            try:
                next_li = driver.find_element(By.XPATH, "//li[a[contains(text(),'بعدی')]]")
                if "disabled" in next_li.get_attribute("class"):
                    print("🚫 We at the last page.")
                    break
                else:
                    next_a = next_li.find_element(By.TAG_NAME, "a")
                    next_url = next_a.get_attribute("href")
                    print(f"➡️ Go to the next page: {next_url}")
                    driver.get(next_url)
                    page_num += 1
                    continue
            except:
                print("Finish")
                break

    finally:
        driver.quit()

    # ذخیره در اکسل با openpyxl
    wb = Workbook()
    ws = wb.active
    ws.title = "Links"
    ws.append(["لینک‌ها"])

    for href in all_links:
        ws.append([href])

    wb.save(out_xlsx)
    print(f"\n📊 Link fined:\n {len(all_links)} saved at the: \n{out_xlsx}")


if __name__ == "__main__":
    start_url = "https://hub.23055.ir/search-lawyer?p_p_id=NetFormRecordsViewer_WAR_NetForm_INSTANCE_KZ6yl1GcLuD5"  # ← URL صفحه اول
    extract_all_pages(start_url, out_xlsx="lawyer_links.xlsx")
