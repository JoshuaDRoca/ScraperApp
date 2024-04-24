from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common import NoSuchElementException, ElementNotInteractableException
from selenium.webdriver import Keys, ActionChains
from selenium.webdriver.common.by import By
from copy import deepcopy
from flask import Flask, render_template, send_file
from io import BytesIO

import openpyxl
import os
import time


def create_app(test_config=None):
    # Create and configure the app
    app = Flask(__name__, instance_relative_config=True)
    app.config.from_mapping(
        SECRET_KEY='dev',
        DATABASE=os.path.join(app.instance_path, 'flaskr.sqlite'),
    )

    if test_config is None:
        # load the instance config, if it exists, when not testing
        app.config.from_pyfile('config.py', silent=True)
    else:
        # load the test config if passed in
        app.config.from_mapping(test_config)

    # ensure the instance folder exists
    try:
        os.makedirs(app.instance_path)
    except OSError:
        pass

    # a simple page that says hello
    @app.route('/hello')
    def hello():
        return 'Hello, World!'

    @app.route('/')
    def index():
        return render_template('index.html')

    @app.route('/start')
    def start():
        wb = scrape()

        file_stream = BytesIO()
        wb.save(file_stream)
        file_stream.seek(0)

        return send_file(file_stream, download_name="Student_Record.xlsx", as_attachment=True)

    def scrape():
        # Instantiating errors
        errors = [NoSuchElementException, ElementNotInteractableException]

        # Instantiating workbook for Excel
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Student Class History"

        # Instantiating and setting up WebDriver
        driver = webdriver.Chrome()
        driver.get("https://degreeworks.clayton.edu/ResponsiveDashboard")

        # creating wait object
        wait = WebDriverWait(driver, timeout=3600, poll_frequency=.2, ignored_exceptions=errors)

        # Waiting until inside DegreeWorks
        wait.until(lambda d: driver.find_element(By.CLASS_NAME, "dashboard-body").is_displayed())

        # Internal check
        print("accessed successfully")

        # Ensuring connection
        driver.implicitly_wait(5)  # DOES NOT PAUSE SCREEN

        # Wait until menu button comes up.
        wait.until(lambda d: driver.find_element(By.XPATH,
                                                 '//*[@id="root"]/div/div[2]/div/main/div/div[1]/div/div/div/div[1]/div/button').is_displayed())

        # Find Student Schedule Page
        menu_button = driver.find_element(By.XPATH,
                                          '//*[@id="root"]/div/div[2]/div/main/div/div[1]/div/div/div/div[1]/div/button')
        ActionChains(driver) \
            .move_to_element(menu_button) \
            .pause(1) \
            .click() \
            .pause(1) \
            .perform()
        class_history = driver.find_element(By.XPATH,
                                            '//*[@id="root"]/div/div[2]/div/main/div/div[1]/div/div/div/div[2]/div/ul/li[2]')
        ActionChains(driver) \
            .move_to_element(class_history) \
            .pause(1) \
            .click() \
            .pause(1) \
            .perform()

        # Loop through every class on schedule and save into array.
        classes = driver.find_elements(By.TAG_NAME, "td")
        class_data = [['Course', 'Titles', 'Grade', 'Credit Hours']]
        x = 0
        temp_data = []
        for i in classes:
            if i.text:
                temp_data.append(str(i.text))
                # print("After appending i.text\n", temp_data)  # internal check
                x += 1
                if x > 3:
                    mover = deepcopy(temp_data)  # brute force solution
                    class_data.append(mover)
                    # print(class_data)  # internal check
                    temp_data.clear()
                    x = 0

        # print(class_data)

        # Printing array to Excel.
        for row in class_data:
            sheet.append(row)

        # Close and save.
        driver.quit()
        # workbook.save("/Users/joshroca/Documents/ScraperApp/scraper/downloads/Student_Record.xlsx")
        # time.sleep(5)

        # Print success message
        # print('Excel file saved successfully')

        return workbook

    return app
