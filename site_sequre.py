# Script to automatically go through the Queen's SeQure app confirmation survey that you don't have COVID..
# So, don't use this if you have COVID symptoms because you will be lying.

# Created by Austin Greisman - 14amfg@queensu.ca
# Date of creation: Aug 16th, 2021.
# V1.0

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.action_chains import ActionChains
firefox_options = Options()
firefox_options.add_argument("--headless")
firefox_options.add_argument('--disable-extensions')
firefox_options.add_argument("--log-level=3")
import time

#Parse for Code
from bs4 import BeautifulSoup

# Send image
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from PIL import Image
from pathlib import Path

carriers = {
	'att':    '@mms.att.net',
	'tmobile':' @tmomail.net',
	'verizon':  '@vtext.com',
	'sprint':   '@page.nextel.com',
    'koodo':   '@msg.koodomobile.com',
	'rogers':  '@sms.rogers.com',
	'telus':    '@msg.telus.com',
	'fido':		'@fido.ca',
	'bell':		'@txt.bell.ca'
}

# For debugging
# python3 -m debugpy --wait-for-client --listen 0.0.0.0:5678 site_sequre.py 

def booker(username, password, phone_num, building="Mitchell Hall", time_on_campus="Morning and Afternoon"):
    '''Inputs:
        username: NETID
        password: Password for NETID Account
        phone_num: Your phone number
        building: Which building you're going to access
        time_on_campus: The time you're planning to be on campus
       Outputs:
        code: Confirmation code for Microsoft Form'''
    try:
        driver = webdriver.Firefox(options=firefox_options)
        # driver = webdriver.Firefox() # For debug and testing.
    except:
        print("Driver wouldn't start. Trying again in a few seconds.")
        time.sleep(10)
        driver = webdriver.Firefox(options=firefox_options)
    
    sequre_url = 'https://queensu.apparmor.com/WebApp/default.aspx?menu=Start+Self+Assessment'
    # In case there is another link added later...
    url_to_load = sequre_url

    # Try out the url.
    driver.set_page_load_timeout(10) # Timeout 10 seconds
    try:
        driver.get(url_to_load)
    except:
        print("System is taking awhile. Will sleep for a few seconds.")
        time.sleep(3)
        driver.get(url_to_load)
    
    # Select Student button
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_Student_4")))
    driver.find_element_by_id("ContentPlaceHolder1_Student_4").click()

    # Log in 
    # Username
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "i0116")))
    driver.find_element_by_id("i0116").send_keys(f'{username}@queensu.ca')
    # Next
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "idSIButton9")))
    actions = ActionChains(driver)
    driver.find_element_by_id("idSIButton9").click()
    time.sleep(3)
    # Password
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "i0118")))
    driver.find_element_by_id("i0118").send_keys(password)
    # Click Sign In
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "idSIButton9")))
    actions = ActionChains(driver)
    actions.move_to_element(driver.find_element_by_id("idSIButton9")).click().perform()

    element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "KmsiDescription")))
    driver.find_element_by_id("idSIButton9").click()

    # Enter Phone number
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_Phone Number")))
    driver.find_element_by_id("ContentPlaceHolder1_Phone Number").send_keys(phone_num)
    # Continue
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_Continue_10")))
    driver.find_element_by_id("ContentPlaceHolder1_Continue_10").click()

    # Click consent check box
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_Consent Checkbox")))
    driver.find_element_by_id("ContentPlaceHolder1_Consent Checkbox").click()
    # Continue
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_Continue_4")))
    driver.find_element_by_id("ContentPlaceHolder1_Continue_4").click()

    # Another consent thing
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_Consent Checkbox 2")))
    driver.find_element_by_id("ContentPlaceHolder1_Consent Checkbox 2").click()

    # Continue
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_Continue_8")))
    driver.find_element_by_id("ContentPlaceHolder1_Continue_8").click()

    # Start selecting No to everything
    select = Select(driver.find_element_by_id('ContentPlaceHolder1_Question 1'))
    select.select_by_value("No")

    select = Select(driver.find_element_by_id('ContentPlaceHolder1_Question 2'))
    select.select_by_value("No")

    select = Select(driver.find_element_by_id('ContentPlaceHolder1_Question 3'))
    select.select_by_value("No")

    select = Select(driver.find_element_by_id('ContentPlaceHolder1_Question 4'))
    select.select_by_value("No")

    select = Select(driver.find_element_by_id('ContentPlaceHolder1_Question 5'))
    select.select_by_value("No")

    # When are you on campus? Error might appear if wrong input.
    select = Select(driver.find_element_by_id('ContentPlaceHolder1_Question 6'))
    select.select_by_value(time_on_campus)
    
    # Where are you on campus? 
    select = Select(driver.find_element_by_id('ContentPlaceHolder1_First Building'))
    select.select_by_value(building)

    # Could add more buildings here...
    
    # Continue...
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_Continue_20")))
    driver.find_element_by_id("ContentPlaceHolder1_Continue_20").click()

    # You don't have COVID. Confirm...
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_Continue_27")))
    driver.find_element_by_id("ContentPlaceHolder1_Continue_27").click()
    
    # Done! Continue
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_lnkButton_2")))
    driver.find_element_by_id("ContentPlaceHolder1_lnkButton_2").click()

    # View Self Assessment
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_My Self Assessment Status_2")))
    driver.find_element_by_id("ContentPlaceHolder1_My Self Assessment Status_2").click()

    # Find window size
    driver.set_window_size(1920, 1080)
    driver.save_screenshot("screenshot.png")
    
    # Time to parse with Beautiful Soup...
    html_soup = BeautifulSoup(driver.page_source, 'html.parser')
    id_index = html_soup.text.find("ID:")
    code = html_soup.text[id_index + 4: id_index + 16]
    print(f'Code for booking generated: {code}')
    driver.quit()
    return code

def CropImage(ImageFileName):
    '''Inputs:
        ImageFilePath: Image for cropping from "booker"
    Ouputs:
        None'''
    img = Image.open(ImageFileName)

    width, height = img.size
    
    # Setting the points for cropped image
    left = width // 2.5
    top = height // 2.2
    right = 1.8* width // 3
    bottom = 3.2 * height // 4

    img_crop = img.crop((left, top, right, bottom))
    # Overwrite previous image.
    img_crop.save("screenshot_Cropped.png")


def SendMail(code, username, password, number, carrier='koodo'):
    '''Inputs:
        code: The code needed to give to Ramzi's new form
        username: Your NETID
        Password: Your Queen\'s Password
        number: Your phone number for testing text to
        carrier: Phone carrier. Needed to generate txt email
    Output:
        None'''
    ImgFileName = Path("screenshot_Cropped.png")
    with open(ImgFileName, 'rb') as f:
        img_data = f.read()

    # Create your Phone number email.
    to_number = str(number) + '{}'.format(carriers[carrier])

    auth = (f'{username}@queensu.ca', password)

    msg = MIMEMultipart()
    msg['Subject'] = ''
    msg['From'] = auth[0]
    msg['To'] = to_number

    # Craft the message
    text = MIMEText(code)
    msg.attach(text)
    image = MIMEImage(img_data, name=os.path.basename(ImgFileName))
    msg.attach(image)

    # Send message
    s = smtplib.SMTP('smtp.office365.com',587)
    s.ehlo()
    s.starttls()
    s.ehlo()
    
    s.login(auth[0], auth[1])
    s.sendmail(auth[0], to_number, msg.as_string())
    s.quit()
    
if __name__ == '__main__':
    # More times and buildings could be added. Just based on our group, I thought this was a good start. The naming convention is the same as in the app, so you could add your own.
    buildings = ['Athletics and Rec Centre (ARC) - Sports Fields',
                'Mitchell Hall', 
                "Jackson Hall"]
    time_on_campus = ["Morning Only", 
                      "Morning and Afternoon", 
                      "Morning and Evening", 
                      "Afternoon Only", 
                      "Afternoon and Evening", 
                      "Evening Only"]
    
    # Parameters to change!
    # ----------
    uname           = 'NETID'           #NETID
    pswrd           = 'PASS'            #Password
    phone_number    = 'XXXXXXXXX'       #No Spaces or dashes
    phone_carrier   = 'Carrier'         #Phone Carrier for sending text. Check the 'carriers' dictionary to see naming.
    building = buildings[1]             #Building you're going to access
    book_time = time_on_campus[0]       #When you're going to be on campus. The script doesn't assume now.
    ## ---------


    # Actually does all the confirmation stuff.
    code = booker(uname, pswrd, phone_number, building, book_time)
    ImageFileName = Path("screenshot.png")
    # Image has a lot of un-need stuff...
    CropImage(ImageFileName)
    # Sends you the code over text and the confirmation QRCode.
    SendMail(code, uname, pswrd, number=int(phone_number), carrier=phone_carrier)
    print("Text Sent. Done!")
    