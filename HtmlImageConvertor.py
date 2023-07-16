from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time
from PIL import Image
import io


#used to crop the HTML screenshots to cut off the scroller 
def cropImage(width,height,imageObj):
    left = 50
    top = 0
    right = width-left
    bottom = height
    image_cropped = imageObj.crop((left, top, right, bottom))
    if(width==400):
        image_cropped.save("CroppedMobile.png")
        print("Mobile cropped render has been saved..")
    else:
        image_cropped.save("CroppedDesktop.png")
        print("Desktop/IPAD cropped render has been saved..")


#A function to set the vieport width to 400px so that mobile version of the html will be renderered
def getMobScreenshot(driver,url):
    currentWidth=400
    driver.set_window_size(currentWidth,800)
    total_height = driver.execute_script("return document.documentElement.scrollHeight")
    driver.set_window_size(currentWidth, total_height)
    time.sleep(2)
    screenshot = driver.get_screenshot_as_png()
    screenshot = Image.open(io.BytesIO(screenshot))
    cropImage(currentWidth,total_height,screenshot)
 
 
#A function to set the vieport width to 700px so that Deskitop / IPAD version of the html will be renderered
def getDeskScreenshot(driver,url):
    currentWidth=700
    driver.set_window_size(currentWidth,800)
    total_height = driver.execute_script("return document.documentElement.scrollHeight")
    driver.set_window_size(currentWidth, total_height)
    time.sleep(2)
    screenshot=driver.get_screenshot_as_png()
    screenshot = Image.open(io.BytesIO(screenshot))
    cropImage(currentWidth,total_height,screenshot)



url = 'open the html and paster the URL of the live HTML here'
options = Options()
options.add_argument('--headless')
options.add_argument('--start-maximized')
driver = webdriver.Chrome(options=options)
driver.get(url)
getMobScreenshot(driver,url)
getDeskScreenshot(driver,url)
driver.quit()