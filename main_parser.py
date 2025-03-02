import schedule
import time
from datetime import datetime
import random
from openpyxl import load_workbook
import requests
from proxy_info import proxies
import pandas as pd

file_path = 'Цены ВБ.xlsx'
main_df = pd.read_excel(file_path, skiprows=2)


def get_wildberries_price(product_id):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
        "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
        "Referer": "https://www.wildberries.ru/",
        "Connection": "keep-alive"
    }

    api_url = f"https://card.wb.ru/cards/detail?appType=1&curr=rub&dest=-1257786&spp=0&nm={product_id}"

    try:
        response = requests.get(api_url, headers=headers, proxies=random.choice(proxies), timeout=10)
        response.raise_for_status()
        data = response.json()

        if "data" in data and "products" in data["data"] and data["data"]["products"]:
            price = data["data"]["products"][0].get("salePriceU", 0) // 100
            return price
        else:
            return "Цена не найдена"
    except requests.RequestException as e:
        print("Ошибка при получении данных:", e)
        return None


def main():
    current_time = datetime.now().replace(minute=0, second=0, microsecond=0)  
    print(f"Функция запущена в: {current_time}")
    prices = []
    
    formatted_time = current_time.strftime("%Y-%m-%d %H:%M:%S")
    
    if datetime.strptime(main_df.iloc[-1, 0], '%Y-%m-%d %H:%M:%S') != \
            datetime.strptime(formatted_time, '%Y-%m-%d %H:%M:%S'):
        
        for product_id in main_df.columns[1:]:  
            price = get_wildberries_price(product_id)

            
            prices.append(price)

        
        new_row = [formatted_time] + prices
        print(f"Новая строка добавлена: {new_row}")

        
        wb = load_workbook(file_path)
        ws = wb.active
        next_row = ws.max_row + 1

        
        for col, value in enumerate(new_row, 1):
            ws.cell(row=next_row, column=col, value=value)

        
        wb.save(file_path)

    else:
        print(f"Строка с датой {formatted_time} уже существует.")



schedule.every().hour.at(":01").do(main)  


while True:
    schedule.run_pending()
    time.sleep(600)  
