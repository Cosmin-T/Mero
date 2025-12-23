# logic/save_cookies.py

from util import *
import json

logger = setup_logging()


def save_cookies_json(page):
    try:
        cookies = page.context.cookies()
        logging.info(f"Cookies: {cookies}")
        print("Cookies:")
        for i, cookie in enumerate(cookies):
            print(f"Cookie {i + 1}:")
            print(
                f"Name: {cookie['name']}, Value: {cookie['value']}, Domain: {cookie['domain']}, Path: {cookie['path']}, Expires: {cookie.get('expires', 'Session')}"
            )
            print("-----------------")
        with open(JSON_PATH, mode="w") as json_file:
            json.dump(cookies, json_file, indent=4)
        logging.info(f"Cookies saved to {JSON_PATH}")
    except Exception as e:
        logging.error(f"Error saving cookies: {e}")
        print(f"Error saving cookies: {e}")


def delete_cookies():
    if os.path.exists(JSON_PATH):
        os.remove(JSON_PATH)
