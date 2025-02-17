import time
import pymysql
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from openpyxl import load_workbook
import XLUtils

driver = webdriver.Chrome()
driver.maximize_window()
driver.get("https://www.moneycontrol.com/fixed-income/calculator/state-bank-of-india-sbi/fixed-deposit-calculator-SBI-BSB001.html")
time.sleep(5)

file = r"data drivern .xlsx"
workbook = load_workbook(file)
sheet_name = "Sheet1"

try:
    con = pymysql.connect(
        host="localhost",
        port=3306,
        user="root",
        passwd="amar@123",
        database="youtube"
    )
    curs = con.cursor()
    # print(curs)
    curs.execute("select * from cal_data")
    rows = curs.fetchall()

    row_index = 2
    for row in rows:
        print(row[0], row[1], row[2], row[3], row[4], row[5])
        principle = row[0]
        rate_of_interest = row[1]
        tenure = row[2]
        tenure_period = row[3]
        frequency = row[4]
        expected_maturity = row[5]

        driver.find_element(By.XPATH, "//input[@name='principal']").clear()
        driver.find_element(By.XPATH, "//input[@name='principal']").send_keys(principle)
        driver.find_element(By.XPATH, "//input[@name='interest']").clear()
        driver.find_element(By.XPATH, "//input[@name='interest']").send_keys(rate_of_interest)
        driver.find_element(By.XPATH, "//input[@name='tenure']").clear()
        driver.find_element(By.XPATH, "//input[@name='tenure']").send_keys(tenure)

        tenure_dropdown = Select(driver.find_element(By.XPATH, "//select[@id='tenurePeriod']"))
        tenure_dropdown.select_by_visible_text(tenure_period)

        frequency_dropdown = Select(driver.find_element(By.XPATH, "//select[@id='frequency']"))
        frequency_dropdown.select_by_visible_text(frequency)

        driver.find_element(By.XPATH, "//a[@href=\"javascript:void(0);\"]").click()
        time.sleep(2)

        maturity_value = driver.find_element(By.XPATH, "//span[@id='resp_matval']").text

        if float(maturity_value) == float(expected_maturity):
            XLUtils.writeData(file, sheet_name, row_index, 8, "Pass")
            XLUtils.fillGreenColor(file,sheet_name,row_index,8)
        else:
            XLUtils.writeData(file, sheet_name, row_index, 8, "Fail")
            XLUtils.fillRedColor(file,sheet_name,row_index,8)

        row_index+=1
except Exception as e:
    print("An Error occurred:",e)
finally:
    driver.quit()
    if 'con' in locals() and con.open:
        con.close()
print("Process completed!")
