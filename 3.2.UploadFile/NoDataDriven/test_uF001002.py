# Generated by Selenium IDE
import pytest
import time
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities

class TestUF001002():
  def setup_method(self, method):
    self.driver = webdriver.Chrome()
    self.vars = {}
  
  def teardown_method(self, method):
    self.driver.quit()
  
  def test_uF001002(self):
    self.driver.get("https://school.moodledemo.net/user/files.php")
    self.driver.set_window_size(1470, 847)
    self.driver.find_element(By.CSS_SELECTOR, ".fa-file-o").click()
    self.driver.find_element(By.NAME, "repo_upload_file").click()
    self.driver.find_element(By.NAME, "repo_upload_file").send_keys("/Users/duyanhle/Desktop/SoftwareTesting-Assignment3/UploadFile/sample-file/one-byte.txt")
    self.driver.find_element(By.XPATH, "//button[contains(.,\'Upload this file\')]").click()
  
