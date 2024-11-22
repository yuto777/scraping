#エクセルで実行したら文字化けしたので修正。
import time
import csv
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# WebDriverの設定
chrome_options = Options()
# chrome_options.add_argument('--headless')  # 必要に応じてコメントアウト
chrome_options.add_argument('--disable-gpu')
driver_path = r'.\chromedriver\chromedriver.exe'  # ChromeDriverのパスを指定

# WebDriverの起動
driver = webdriver.Chrome(service=Service(driver_path), options=chrome_options)

# MonotaroのURL
url = "https://www.monotaro.com/s/c-20/"

# 特殊文字を置換する関数
def replace_invalid_chars(text, encoding="utf-8"):
    try:
        text.encode(encoding)  # エンコード可能か試す
        return text
    except UnicodeEncodeError:
        return text.replace("\uff5e", "～").replace("\u3000", " ")  # 一部の特殊文字を変換

try:
    # ページを開く
    print("ページを開いています...")
    driver.get(url)

    # ページ全体が読み込まれるまで待機
    print("ページ全体の読み込みを待機...")
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))

    # 商品リンクが表示されるまで待機
    print("商品情報を待機しています...")
    wait = WebDriverWait(driver, 20)
    product_elements = wait.until(
        EC.presence_of_all_elements_located(
            (By.CSS_SELECTOR, ".TextLink.u-FontSize--Lg.SearchResultImage__TextLink")
        )
    )
    price_elements = wait.until(
        EC.presence_of_all_elements_located(
            (By.CSS_SELECTOR, ".Price.Price--Md")
        )
    )

    # 商品名、値段、リンクを取得
    print("商品情報を取得しています...")
    products = []
    for product_elem, price_elem in zip(product_elements, price_elements):
        product_name = replace_invalid_chars(product_elem.text)
        product_link = "https://www.monotaro.com" + product_elem.get_attribute("href")  # フルURL化
        product_price = replace_invalid_chars(price_elem.text.strip().replace("\n", " "))  # 価格情報を整形
        products.append((product_name, product_price, product_link))

    # CSVに保存（UTF-8 with BOM）
    print("CSVに保存中...")
    csv_file = "monotaro_products.csv"
    with open(csv_file, mode="w", encoding="utf-8-sig", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(["商品名", "価格", "URL"])  # ヘッダー行
        writer.writerows(products)  # データ行を追加

    print(f"CSVファイル '{csv_file}' に保存しました！")

except Exception as e:
    print(f"エラーが発生しました: {e}")
    import traceback
    traceback.print_exc()  # 詳細なエラーログを表示
finally:
    # WebDriverを閉じる
    driver.quit()
