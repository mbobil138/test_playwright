import asyncio
from playwright.async_api import async_playwright
import openpyxl

async def scrape():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        page = await browser.new_page()
        await page.goto('https://rozetka.com.ua/')

        data = {}

        # Введення запиту в пошук
        await page.fill("//input[contains(@class, 'search')]", 'Apple iPhone 15 128GB Black')
        await page.click("//button[contains(@class, 'search-form__submit')]")
        await page.wait_for_timeout(2000)

        # Перехід на сторінку товару
        await page.click('(//rz-indexed-link[@class="product-link"])[1]')
        await page.wait_for_timeout(2000)

        # Отримання інформації про товар
        data['Назва'] = await page.inner_text('//h1[@class="title__font"]')

        color = await page.inner_text('(//div[@class="var-options"]/p[@class="text-base mb-2"])[1]')
        data['Колір'] = color.strip()

        memory = await page.inner_text('(//div[@class="var-options"]/p[@class="text-base mb-2"])[2]')
        data["Пам'ять"] = memory.strip()

        seller = await page.get_attribute('//img[@alt="Rozetka"]', 'alt')
        data['Продавець'] = seller.strip()

        try:
            price = await page.inner_text('//p[@class="product-price__small"]')
            discount = await page.inner_text('//p[contains(@class, "product-price__big")]')
            data['Ціна'] = price.strip()
            data['Знижка'] = discount.strip()
        except:
            data['Ціна'] = 'Не відображає ціну'
            data['Знижка'] = 'No Discount'

        # Фото
        try:
            photo_elements = await page.locator('//ul[@class="simple-slider__list"]//img').all()
            photo_urls = [await img.get_attribute('src') for img in photo_elements]
            data['Фото'] = '\n'.join(photo_urls)
        except:
            data['Фото'] = 'Не знайшло правильний лист'

        code = await page.inner_text('(//div[@class="rating text-base"])[2]')
        data['Код товару'] = code.strip()

        try:
            reviews_link = await page.get_attribute('//a[contains(@href, "comments")]', 'href')
            data['Відгуки'] = reviews_link if reviews_link else 'No reviews'
        except:
            data['Відгуки'] = 'No reviews'

        series = await page.inner_text('//a[contains(@href, "series=iphone-15")]')
        data['Серія'] = series.strip()

        try:
            diag = await page.inner_text("//span[text()='6.1']")
            data['Діагональ'] = diag.strip()
        except:
            data['Діагональ'] = 'Not Found'

        print(data)

        # Робота з Excel
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.title = "Товари"

        # Додаємо заголовки та дані
        sheet.append(list(data.keys()))
        sheet.append(list(data.values()))

        wb.save("Result.xlsx")
        print("Дані успішно збережено в Result.xlsx!")

        await browser.close()

asyncio.run(scrape())