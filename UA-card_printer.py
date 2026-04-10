from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd  # 用於處理 Excel 文件
import os
import time
import requests
from PIL import Image, ImageTk
from math import floor
import re
import tkinter as tk
from tkinter import simpledialog, messagebox
input_url = ""

root = tk.Tk()
root.title("UA印牌機器")
root.geometry("900x470")
root.resizable(False, False)
root.attributes("-alpha", 0.0) 

bg_image = Image.open("C:/Users/user/Documents/Card_py/png/image.png")
bg_photo = ImageTk.PhotoImage(bg_image)

canvas = tk.Canvas(root, width=400, height=300)
canvas.pack(fill="both", expand=True)
canvas.create_image(0, 0, image=bg_photo, anchor="nw")


canvas.create_text(450, 90, text="請輸入卡組網址(洛基亞)", fill="white", font=("微軟正黑體", 16, "bold"))
entry = tk.Entry(root, font=("微軟正黑體", 14), width=30)
canvas.create_window(450, 150, window=entry)

def submit():
    global input_url
    input_url = entry.get()
    if input_url:
        root.destroy()
    else:
        messagebox.showwarning("提醒", "請輸入網址")

btn = tk.Button(root, text="確定", font=("微軟正黑體", 12), command=submit)
canvas.create_window(450, 210, window=btn)

def fade_in(alpha=0.0):
    if alpha < 1.0:
        alpha += 0.05
        root.attributes("-alpha", alpha)
        root.after(30, lambda: fade_in(alpha))


fade_in()

root.mainloop()

speace = []  # 用於存儲空格數量的列表
n = 0
# 設定圖片來源資料夾與輸出 PDF 路徑
image_folder = "C:/Users/user/Documents/Card_py/UA_png"
output_pdf_path = "C:/Users/user/Documents/Card_py/PDF"

# 設定 Chrome WebDriver
options = webdriver.ChromeOptions()
#options.add_argument('--headless')  # 如果需要無頭模式，可以取消註解
options.add_argument('--ignore-certificate-errors')
options.add_argument('--allow-insecure-localhost')
os.makedirs("C:/Users/user/Documents/Card_py/UA_png", exist_ok=True)
# MCR 4UA36BT_1009 UA36BT/MCR-1-009  3UA36BT_1009_2 UA36BT/MCR-1-009|3UA36BT_1009

for file_name in os.listdir(image_folder):
    
    if file_name.endswith(".jpg"):  # 找到所有 .txt 檔案
        os.remove(os.path.join(image_folder, file_name))
        print(f"{file_name} 已刪除")

def download_image(img_url, save_path):
    try:
        response = requests.get(img_url, stream=True)
        if response.status_code == 200:
            with open(save_path, 'wb') as file:
                for chunk in response.iter_content(1024):
                    file.write(chunk)
            print(f"圖片已下載：{save_path}")
        else:
            print(f"無法下載圖片，狀態碼：{response.status_code}")
    except Exception as e:
        print(f"下載圖片時發生錯誤：{e}")



delay_time = 30  # 等待時間（秒）

# 用於存儲卡片數據的列表
card_data = []
card_name = []
card_number = []
is_blood_card = []
card_first = []

# 解析版本代碼
version_match = re.search(r"Version=([A-Z0-9]+)", input_url)
version = version_match.group(1) if version_match else "未知"
print("版本代碼：", version)

# 擷取 Deck 的部分
deck_str = input_url.split("Deck=")[-1]

# 拆分每張卡片的資訊
card_entries = deck_str.split("|")

for entry in card_entries:
    match = re.match(r"(\d)([A-Z]+)(\d*[A-Z]*)_(\d{4})(_\d)?", entry)
    if match:
        quantity = int(match.group(1))                  # 張數
        name = match.group(2) + match.group(3)          # 套件名（例如 UA36BT 或 UAPR）
        number = match.group(4)                         # 卡號（例如 1009）
        suffix = match.group(5)                         # 是否有 _2 等附加資料

        card_name.append(name)
        card_number.append(number)
        speace.append(quantity)
        is_blood_card.append(1 if suffix == "_2" else 2 if suffix == "_3" else 0)

card_first = [number[0] for number in card_number]
card_number = [i[1:] for i in card_number]
for i in range(0, len(card_name)):
    card_data.append(f"{card_name[i]}/{version}-{card_first[i]}-{card_number[i]}")
    
print("卡片名稱：", card_data)
print("卡片數量：", speace)

# 開啟瀏覽器並進行操作
img_ulr = "https://www.unionarena-tcg.com/jp/cardlist/?search=true"
driver = webdriver.Chrome(service=Service("C:/chromedriver-win64/chromedriver.exe"), options=options)
driver.get(img_ulr)

try:
    # 等待頁面加載完成
    cardlist = WebDriverWait(driver, delay_time).until(
        EC.presence_of_element_located((By.TAG_NAME, "body"))
    )
    main = cardlist.find_element(By.TAG_NAME, "main")
    article = main.find_element(By.TAG_NAME, "article")
    card_marp = WebDriverWait(article, delay_time).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, ".cardMainWrap.fadein.is-fadein"))
    )
    form = card_marp.find_element(By.TAG_NAME, "form")
    primarySearchWrap = form.find_element(By.CLASS_NAME, "primarySearchWrap")
    aside = primarySearchWrap.find_element(By.TAG_NAME, "aside")
    priSearchInner = aside.find_element(By.CLASS_NAME, "priSearchInner")
    freewordsCol = priSearchInner.find_element(By.CLASS_NAME, "freewordsCol")
    freewordsTxtCol = freewordsCol.find_element(By.CLASS_NAME, "freewordsTxtCol")
    input = WebDriverWait(freewordsTxtCol, delay_time).until(
        EC.presence_of_element_located((By.CLASS_NAME, "freewords"))
    )


    submitBtn = aside.find_element(By.CLASS_NAME, "submitBtn")
    button = submitBtn.find_element(By.TAG_NAME, "input")

    input.send_keys(card_data[0])
    button.click()
    
    for i in range(1, len(card_data)+1):
        try:
            # 等待頁面重新加載
            cardlist = WebDriverWait(driver, delay_time).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            main = cardlist.find_element(By.TAG_NAME, "main")
            article = main.find_element(By.TAG_NAME, "article")
            card_marp = WebDriverWait(article, delay_time).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, ".cardMainWrap.fadein.is-fadein"))
            )
            form = card_marp.find_element(By.TAG_NAME, "form")
            primarySearchWrap = form.find_element(By.CLASS_NAME, "primarySearchWrap")
            aside = primarySearchWrap.find_element(By.TAG_NAME, "aside")
            priSearchInner = aside.find_element(By.CLASS_NAME, "priSearchInner")
            freewordsCol = priSearchInner.find_element(By.CLASS_NAME, "freewordsCol")
            freewordsTxtCol = freewordsCol.find_element(By.CLASS_NAME, "freewords")
            submitBtn = WebDriverWait(aside, delay_time).until(
                EC.presence_of_element_located((By.CLASS_NAME, "submitBtn"))
            )
            button = submitBtn.find_element(By.TAG_NAME, "input")

            # 重新定位圖片元素
            cardContentsCol = card_marp.find_element(By.CLASS_NAME, "cardContentsCol")
            cardlistWrap = cardContentsCol.find_element(By.CLASS_NAME, "cardlistWrap")
            cardlistCol = cardlistWrap.find_element(By.CLASS_NAME, "cardlistCol")
            li = cardlistCol.find_elements(By.TAG_NAME, "li")
            if is_blood_card[i-1] == 1 and len(li) > 1:
                li = li[1]
            elif is_blood_card[i-1] == 2 and len(li) > 1:
                li = li[2]
            else:
                li = li[0]
            a = li.find_element(By.TAG_NAME, "a")
            driver.execute_script("arguments[0].scrollIntoView();", a)
            img = WebDriverWait(a, delay_time).until(
                EC.presence_of_element_located((By.TAG_NAME, "img"))
            )

            # 等待圖片的 src 屬性更新
            WebDriverWait(driver, delay_time).until(
                lambda d: img.get_attribute("src") and "dummy.gif" not in img.get_attribute("src")
            )

            img_url = img.get_attribute("src")  # 取得圖片 URL
            print(f"圖片 URL：{img_url}")

            # 檢查是否為有效圖片 URL
            if "dummy.gif" in img_url:
                print(f"跳過無效圖片：{img_url}")
                continue

            save_path = os.path.join("C:/Users/user/Documents/Card_py/UA_png", f"card_{i}.jpg")
            download_image(img_url, save_path)

            # 清除搜尋框並輸入下一張卡片名稱
            freewordsTxtCol.clear()
            freewordsTxtCol.send_keys(card_data[i])
            button.click()

        except Exception as e:
            print(f"下載卡片 {card_data[i]} 時發生錯誤：{e}")
            continue




except Exception as e:
    print(f"發生錯誤：{e}")
finally:
    # 關閉瀏覽器
    driver.quit()


A4_WIDTH_CM = 21.0
A4_HEIGHT_CM = 29.7
DPI = 300

# 卡片尺寸（你提供的：6.47cm x 9.02cm）
CARD_WIDTH_CM = 6.47
CARD_HEIGHT_CM = 9.02

# 換算為像素
a4_width_px = int(A4_WIDTH_CM / 2.54 * DPI)
a4_height_px = int(A4_HEIGHT_CM / 2.54 * DPI)
card_width_px = int(CARD_WIDTH_CM / 2.54 * DPI)
card_height_px = int(CARD_HEIGHT_CM / 2.54 * DPI)


cols = floor(a4_width_px / card_width_px)
rows = floor(a4_height_px / card_height_px)
cards_per_page = cols * rows

images = []
e_count = 0
# 排序圖片檔案名稱，按數字順序
image_files = sorted(
    [f for f in os.listdir(image_folder) if f.lower().endswith(('.png', '.jpg', '.jpeg'))],
    key=lambda x: int(re.search(r'\d+', x).group()) if re.search(r'\d+', x) else float('inf')
)

for i, count in enumerate(speace):
    if i >= len(image_files):
        break
    img_path = os.path.join(image_folder, image_files[i])
    img = Image.open(img_path).convert("RGB")
    e_count = count + e_count

    
    img = img.resize((card_width_px, card_height_px), Image.LANCZOS)

    # 根據 count 數量複製卡片
    for _ in range(count):
        images.append(img.copy())

# === 將圖片排列成多頁 PDF ===
pages = []
for i in range(0, len(images), cards_per_page):
    page_img = Image.new("RGB", (a4_width_px, a4_height_px), "white")
    for index_on_page, img in enumerate(images[i:i + cards_per_page]):
        row = index_on_page // cols
        col = index_on_page % cols
        x = col * card_width_px
        y = row * card_height_px
        page_img.paste(img, (x, y))
    pages.append(page_img)

# === 輸出 PDF ===
if pages:
    pages[0].save(output_pdf_path, save_all=True, append_images=pages[1:])
    print(f"✅ PDF 已儲存：{output_pdf_path}")
else:
    print("⚠️ 沒有圖片可處理。") 
