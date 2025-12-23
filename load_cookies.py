# logic/load_cookies.py

from util import *
import json


def normalize_cookie(cookie):
    cookie = cookie.copy()

    # First check if sameSite exists and isn't None
    if not cookie.get("sameSite"):
        cookie["sameSite"] = "Lax"
    else:
        # Only try to lowercase if it's a string
        sameSite = cookie["sameSite"]
        if isinstance(sameSite, str):
            if sameSite.lower() == "strict":
                cookie["sameSite"] = "Strict"
            elif sameSite.lower() == "lax":
                cookie["sameSite"] = "Lax"
            elif sameSite.lower() == "none":
                cookie["sameSite"] = "None"
        else:
            cookie["sameSite"] = "Lax"

    return cookie


def load_cookies_json(context):
    try:
        with open(JSON_PATH, "r") as f:
            cookies = json.load(f)

        if cookies:  # Only process cookies if the file has content
            normalized_cookies = [normalize_cookie(cookie) for cookie in cookies]
            context.add_cookies(normalized_cookies)
            logger.info(f"Successfully loaded {len(cookies)} cookies from {JSON_PATH}")
        else:
            logger.info(f"Cookie file is empty, continuing without cookies")

    except FileNotFoundError:
        logger.info(f"Cookie file not found at {JSON_PATH}, continuing without cookies")
        return
    except json.JSONDecodeError:
        logger.info(
            f"Invalid or empty JSON in cookie file: {JSON_PATH}, continuing without cookies"
        )
        return
    except Exception as e:
        logger.error(f"Error loading cookies: {str(e)}")
        return
