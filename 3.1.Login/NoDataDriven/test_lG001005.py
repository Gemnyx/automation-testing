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

class TestLG001005():
  def setup_method(self, method):
    self.driver = webdriver.Chrome()
    self.vars = {}
  
  def teardown_method(self, method):
    self.driver.quit()
  
  def test_lG001005(self):
    self.driver.get("https://school.moodledemo.net/login/index.php")
    self.driver.set_window_size(1470, 847)
    self.driver.find_element(By.CSS_SELECTOR, ".login-wrapper").click()
    self.driver.find_element(By.ID, "username").send_keys("student0")
    self.driver.find_element(By.ID, "password").send_keys("")
    self.driver.find_element(By.ID, "loginbtn").click()
    elements = self.driver.find_elements(By.CSS_SELECTOR, ".alert")
    assert len(elements) > 0
  
