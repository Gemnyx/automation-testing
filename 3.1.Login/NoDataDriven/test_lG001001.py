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

class TestLG001001():
  def setup_method(self, method):
    self.driver = webdriver.Chrome()
    self.vars = {}
  
  def teardown_method(self, method):
    self.driver.quit()
  
  def test_lG001001(self):
    self.driver.get("https://school.moodledemo.net/login/index.php")
    self.driver.set_window_size(1470, 847)
    self.driver.find_element(By.ID, "username").send_keys("student")
    self.driver.find_element(By.ID, "password").send_keys("moodle")
    self.driver.find_element(By.ID, "loginbtn").click()
    elements = self.driver.find_elements(By.ID, "user-menu-toggle")
    assert len(elements) > 0
    self.driver.find_element(By.ID, "user-menu-toggle").click()
    self.driver.find_element(By.LINK_TEXT, "Log out").click()
  
