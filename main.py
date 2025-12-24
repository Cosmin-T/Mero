# main.py

import re
from cmath import log

import easyocr
import pytesseract
from load_cookies import *
from PIL import Image
from playwright.sync_api import Playwright, sync_playwright
from util import *

logger = setup_logging()


def run(playwright: Playwright) -> None:
    logger.info("\n\n")
    logger.info("Starting login process")
    context = playwright.chromium.launch_persistent_context(
        "./chrome-data",
        headless=False,
        args=[
            "--disable-blink-features=AutomationControlled",
            "--disable-dev-shm-usage",
            "--no-sandbox",
            "--disable-setuid-sandbox",
            "--ignore-certificate-errors",
            # f'--user-agent={random.choice(USER_AGENTS)}'
        ],
        viewport={"width": 3024, "height": 1440},
        java_script_enabled=True,
        bypass_csp=True,
        ignore_https_errors=True,
    )

    context.add_init_script("""
        Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
    """)

    load_cookies_json(context)

    context.set_default_timeout(15000)

    page = context.new_page()
    page.goto(URL)
    page.wait_for_load_state("networkidle")

    login_identifier = (
        page.locator("div")
        .filter(has_text=re.compile(r"^Utilizează numărul de mobil$"))
        .nth(1)
    )

    if login_identifier.is_visible():
        logger.info("Login identifier triggered")
        login_identifier.click()
        page.wait_for_load_state("networkidle")
        page.get_by_role("textbox", name="Introdu nr. tău de telefon").fill(
            "0799945509"
        )
        logger.info("Filled phone number")
        page.get_by_text("Trimite cod verificare").click()
        logger.info("Sent Access Code, waiting...'")
        page.wait_for_load_state("networkidle")

        access_wall_identifier = page.get_by_role("textbox").filter(
            has=page.locator('[inputmode="numeric"]')
        )
        while len(access_wall_identifier.get_attribute("value") or "") != 6:
            time.sleep(1)

        logger.info("Access code received")
        page.locator("div").filter(has_text=re.compile(r"^Verifică cod$")).nth(
            1
        ).click()
        logger.info("Access code verified")
        logger.info("Logged in successfully")
        page.wait_for_load_state("networkidle")

        try:
            page.locator(
                "div:nth-child(4) > div > .css-175oi2r > svg > g > path"
            ).click()
            logger.info("Clicked on 'Zilnic'")
            time.sleep(1)

            page.locator(
                "div:nth-child(2) > .css-175oi2r.r-1i6wzkk > div > div > svg"
            ).click()
            logger.info("Clicked on 'Saptamanal'")
            time.sleep(1)

        except Exception as e:
            logger.error(f"An error occurred: {e}")

    else:
        logger.info("Login identifier not triggered, already logged in")

    page.wait_for_load_state("networkidle")
    # page.pause()

    # page.get_by_text("14:00", exact=True).first.scroll_into_view_if_needed()
    page.evaluate(
        "Array.from(document.querySelectorAll('div.css-146c3p1')).find(el => el.textContent === '10:00').scrollIntoView()"
    )
    logger.info("Scroll to the top of the page")

    # Extract by column using bounding boxes
    schedule_by_day = page.evaluate(r"""
        () => {
            const schedule = {
                'Luni': [],
                'Marți': [],
                'Miercuri': [],
                'Joi': [],
                'Vineri': [],
                'Sâmbătă': [],
                'Duminică': []
            };

            // Get all appointment blocks - they have specific styling
            const appointmentBlocks = Array.from(document.querySelectorAll('div')).filter(el => {
                const style = window.getComputedStyle(el);
                const bgColor = style.backgroundColor;
                // Look for colored blocks (appointments have background colors)
                return bgColor && bgColor !== 'rgba(0, 0, 0, 0)' && bgColor !== 'transparent' &&
                       el.textContent.match(/\d{1,2}:\d{2}\s*[-–]\s*\d{1,2}:\d{2}/);
            });

            // Get day column boundaries
            const dayHeaders = Array.from(document.querySelectorAll('*')).filter(el => {
                const text = el.textContent.trim();
                return Object.keys(schedule).includes(text) && el.children.length === 0;
            });

            const dayBounds = {};
            dayHeaders.forEach(header => {
                const dayName = header.textContent.trim();
                const rect = header.getBoundingClientRect();
                dayBounds[dayName] = {
                    left: rect.left,
                    right: rect.right,
                    center: (rect.left + rect.right) / 2
                };
            });

            // Assign appointments to days based on horizontal position
            appointmentBlocks.forEach(block => {
                const blockRect = block.getBoundingClientRect();
                const blockCenter = (blockRect.left + blockRect.right) / 2;

                let closestDay = null;
                let closestDistance = Infinity;

                Object.entries(dayBounds).forEach(([day, bounds]) => {
                    const distance = Math.abs(blockCenter - bounds.center);
                    if (distance < closestDistance) {
                        closestDistance = distance;
                        closestDay = day;
                    }
                });

                if (closestDay) {
                    const fullText = block.textContent.trim().split('\n').join(' ');
                    schedule[closestDay].push(fullText);
                }
            });

            return schedule;
        }
    """)

    # Build schedule output
    schedule_output = ["\n=== ROXANA'S SCHEDULE ===\n"]

    for day, appointments in schedule_by_day.items():
        schedule_output.append(f"{day}")
        if appointments:
            # Remove duplicates and filter out garbage data
            seen = set()
            for appt in appointments:
                # Skip if too long (likely garbage)
                if len(appt) > 300:
                    continue

                # Extract time range (e.g., "10:00 - 11:30")
                time_match = re.match(
                    r"(\d{1,2}:\d{2}\s*[-–]\s*\d{1,2}:\d{2})(.*)", appt
                )
                if time_match:
                    time_range = time_match.group(1).replace("–", "-")
                    description = time_match.group(2).strip()

                    # Create unique key for deduplication
                    key = f"{time_range}::{description}"
                    if key not in seen:
                        seen.add(key)
                        schedule_output.append(f"{time_range} :: {description}")
        else:
            schedule_output.append("  No appointments")
        schedule_output.append("")

    # Write to file
    with open("schedule.txt", "w", encoding="utf-8") as f:
        f.write("\n".join(schedule_output))

    # Log to logger
    logger.info("\n=== ROXANA'S SCHEDULE ===")
    for line in schedule_output[1:]:  # Skip the header since we already logged it
        if line:
            logger.info(line)

    logger.info("Schedule saved to schedule.txt")

    page.close()
    context.close()


with sync_playwright() as playwright:
    run(playwright)
