<h1 align="center">🛍️ Wildberries Analytics PRO</h1>

<p align="center">
  <img src="[https://img.shields.io/badge/Python-3.10+-blue.svg](https://img.shields.io/badge/Python-3.10+-blue.svg)" alt="Python">
  <img src="[https://img.shields.io/badge/Status-Production_Ready-success.svg](https://img.shields.io/badge/Status-Production_Ready-success.svg)" alt="Status">
  <img src="[https://img.shields.io/badge/API-Mobile_Internal-orange.svg](https://img.shields.io/badge/API-Mobile_Internal-orange.svg)" alt="API">
</p>

## 📋 О проекте
Десктопный инструмент для мгновенного парсинга и глубокой аналитики ниш на Wildberries. 

Программа решает главную боль селлеров и менеджеров маркетплейсов: **«Как быстро оценить конкурентов, объем рынка и тренды, не прокликивая десятки страниц вручную?»**.

В отличие от стандартных парсеров на Selenium, данный скрипт работает напрямую через **внутреннее Mobile API** платформы, что обеспечивает колоссальную скорость сбора данных без нагрузки на систему.

## ✨ Ключевые возможности
- 🚀 **Сверхскорость:** Сбор до 500 товаров занимает всего 3-5 секунд.
- 🛡️ **Stealth Mode:** Встроена система имитации реального пользователя (подмена Headers, Cookies и QueryID) для обхода защиты Cloudflare.
- 💰 **Метрика LTV (Lifetime Value):** Скрипт не просто собирает цены, но и оценивает историческую выручку каждой карточки (`Цена × Кол-во отзывов`).
- 📊 **Smart Dashboard:** Автоматическая генерация второго листа в Excel со сводной аналитикой (средний чек, емкость рынка, ТОП-3 бренда ниши).
- 🎨 **Готовый отчет:** На выходе формируется стилизованный `.xlsx` файл с авто-фильтрами, денежными форматами и кликабельными ссылками.

## 📸 Скриншоты работы

**1. Процесс сбора данных (CLI):**
![Console Run](assets/console_run.jpg)

**2. Дашборд аналитики (Сводка по нише):**
![Dashboard](assets/excel_dashboard.jpg)

**3. Умная таблица с результатами:**
![Excel Data](assets/excel_data.jpg)

## 🛠 Технический стек
- **Язык:** `Python 3.10`
- **Сетевые запросы:** `requests`
- **Анализ данных:** `pandas`
- **Форматирование отчетов:** `openpyxl`
- **Интерфейс:** `colorama`, `tqdm`

## ⚙️ Установка и запуск (Для разработчиков)

1. Склонируйте репозиторий:
   ```bash
   git clone https://github.com/ВАШ_НИК/WB_Analytics_Tool.git
