from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
PATH = "C:\Program Files (x86)\chromedriver.exe"

def webscrape(keyword):
	driver = webdriver.Chrome(PATH)
	driver.get("https://digikala.com")
	search = driver.find_element_by_name("q")
	search.send_keys(keyword)
	search.send_keys(Keys.RETURN)
	products = ""
	try:
	    main = WebDriverWait(driver, 10).until(
	        EC.presence_of_element_located((By.ID, "main"))
	    )
		titles = main.find_elements_by_css_selector("a.js-product-url")
	    for title in titles:
	    	product = title.text
	    	products += product
	    return products
	except:
	    driver.quit()

