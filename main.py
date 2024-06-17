import argparse
import os
from operator import itemgetter

import openpyxl
import requests
from dotenv import load_dotenv

import category_map as cm

load_dotenv()

MAX_IMAGE_SIZE_MB = 6

output_dir = 'output'


def get_access_token():
    return os.getenv('ACCESS_TOKEN')


def get_owner_id():
    return os.getenv('OWNER_ID')


def get_vk_version():
    return os.getenv('VK_VERSION')


def get_vk_market_url():
    return os.getenv('VK_MARKET_URL')


def create_excel_file(file_name, items):
    workbook = openpyxl.Workbook()
    worksheet = workbook.active

    row = 1
    for item in items:
        worksheet[f"A{row}"] = item["id"]
        worksheet[f"B{row}"] = map_category(item["category"]['name'])
        worksheet[f"C{row}"] = item["title"]
        worksheet[f"D{row}"] = item["price"]["amount"][:-2]
        worksheet[f"E{row}"] = item["description"]
        worksheet[f"F{row}"] = get_primary_image_url(item)
        worksheet[f"G{row}"] = get_secondary_image_url(item)
        row += 1

    workbook.save(f"output/{file_name}.xlsx")
    print(f"Данные успешно экспортированы в файл {file_name}.xlsx")


def get_image_url(photos, index):
    if not photos or index >= len(photos):
        return None

    photo = photos[index]
    photo_sizes = sorted(photo["sizes"], key=lambda x: x["height"] * x["width"], reverse=True)

    for size in photo_sizes:
        image_url = size["url"]
        try:
            image_size_mb = len(requests.get(image_url).content) / (1024 * 1024)
            if image_size_mb <= MAX_IMAGE_SIZE_MB:
                return image_url
        except requests.exceptions.RequestException:
            continue

    return None


def get_primary_image_url(item):
    if "photos" not in item:
        return None
    return get_image_url(item["photos"], 0)


def get_secondary_image_url(item):
    if "photos" not in item or len(item["photos"]) < 2:
        return None
    return get_image_url(item["photos"], 1)


def map_category(category):
    return cm.mapping.get(category, category)


def get_vk_market_data(access_token, owner_id, vk_version, vk_market_url):
    params = {
        "owner_id": owner_id,
        "count": 18,
        "access_token": access_token,
        "v": vk_version,
        "extended": 1
    }

    response = requests.get(vk_market_url, params=params)
    data = response.json()

    if "error" in data:
        print(f"Error: {data['error']['error_msg']}")
        return None
    else:
        return data["response"]["items"]


def main():
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    parser = argparse.ArgumentParser(description="Экспорт данных из VK Market в Excel")
    parser.add_argument("--file-name", default="vk_market_data", help="Имя файла для экспорта")

    args = parser.parse_args()

    access_token = get_access_token()
    owner_id = get_owner_id()
    vk_version = get_vk_version()
    vk_market_url = get_vk_market_url()

    items = get_vk_market_data(access_token, owner_id, vk_version, vk_market_url)

    if items:
        items = sorted(items, key=itemgetter("title"))
        create_excel_file(args.file_name, items)


if __name__ == "__main__":
    main()
