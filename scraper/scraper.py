from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import sys, os

# Getting Login Info
try:
    login_username = os.environ['USERNAME']
    login_password = os.environ['PASSWORD']
except:
    print("please set USERNAME and PASSWORD envirnoment values for MyConcordia!")

# Storing jobname
try:
	job_name = sys.argv[1]
except:
	print("please enter the job name argument!")
	sys.exit(0)
    
# Creating webdriver
filedir = os.path.dirname(os.path.abspath(__file__))
chromedriver_path = os.path.join(filedir, "chromedriver")
driver = webdriver.Chrome(executable_path=chromedriver_path)
driver.maximize_window() # window needs to me maximized for the compass sidebar to appear

# MyConcordia login
driver.get("https://my.concordia.ca/psp/upprpr9/?cmd=login&languageCd=ENG")
assert "MyConcordia - Concordia University" in driver.title
main_page = driver.current_window_handle 

username = driver.find_element_by_id("userid")
username.clear()
username.send_keys("a_mbl")

password = driver.find_element_by_name("pwd")
password.clear()
password.send_keys("Shadowhideyou4%")

driver.find_element_by_class_name("form_button_submit").click()

# Navigate to COOP Compass
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "fldra_CU_STUDENT_REQUESTS"))).click()
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.LINK_TEXT, "Co-op COMPASS"))).click()

for handle in driver.window_handles: 
    if handle != main_page: 
        compass_page = handle 
driver.switch_to.window(compass_page)

WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.LINK_TEXT, "Jobs"))).click()
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.LINK_TEXT, "For My Program"))).click()

# Navigate to specific job page
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.LINK_TEXT, job_name))).click()

for handle in driver.window_handles: 
    if handle != main_page and handle != compass_page: 
        job_page = handle 
driver.switch_to.window(job_page)

# Creating temp info file
tempfile = open("temp.txt","w+")

# Gather employer info
table_trs = driver.find_elements_by_xpath('//div[@id="postingDiv"]/div[4]/div[2]/table/tbody/tr')

for i in range(9):
    tr = table_trs[i]
    name_td = tr.find_elements_by_xpath(".//td[1]/strong")
    value_td = tr.find_elements_by_xpath(".//td[2]")
    name = str.strip(name_td[0].get_attribute("innerHTML"))
    value = str.strip(value_td[0].get_attribute("innerHTML"))
    tempfile.write(name + "->" + value + "\n")

#closing tempfile
tempfile.close()