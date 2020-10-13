from selenium import webdriver
import pytest
import time
import pandas as pd


def firefox_driver_init():
    """Init the driver instance for firefox geckodriver."""
    driver = webdriver.Firefox(
        executable_path="../Webdrivers/geckodriver",
    )
    driver.maximize_window()
    driver.get("http://www.rpachallenge.com/")
    return driver

def test_rpa():
    """The goal of this challenge is to create a workflow that will input data from a spreadsheet into the form fields on the screen"""
    driver = firefox_driver_init()
    driver.find_element_by_xpath("//button[text()='Start']").click()
    insert_values(driver)
    driver.close()
    
def insert_values(driver):
    """Inserts in each field on the website based on xlsx values."""
    list_fields = ["labelCompanyName", "labelRole", "labelPhone", "labelAddress", "labelFirstName", "labelLastName", "labelEmail"]
    dict_final = {}
    #Reads the xlsx and convert to a python dict.
    temp_dic = pd.read_excel('challenge.xlsx', index_col=0).to_dict()
    for k,v in temp_dic["Last Name "].items():
        dict_final["labelFirstName"] = k
        dict_final["labelLastName"] = v
        dict_final["labelCompanyName"] = temp_dic["Company Name"][dict_final["labelFirstName"]]
        dict_final["labelRole"] = temp_dic["Role in Company"][dict_final["labelFirstName"]]
        dict_final["labelAddress"] = temp_dic["Address"][dict_final["labelFirstName"]]
        dict_final["labelEmail"] = temp_dic["Email"][dict_final["labelFirstName"]]
        dict_final["labelPhone"] = temp_dic["Phone Number"][dict_final["labelFirstName"]]
        for field in list_fields:
            driver.find_element_by_xpath(f"//input[@ng-reflect-name='{field}']").send_keys(dict_final[field])
        driver.find_element_by_xpath("//input[@value='Submit']").click()
    

        

