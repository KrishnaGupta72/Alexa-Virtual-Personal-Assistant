from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time

#For system speaking
import win32com.client as wincl
speak = wincl.Dispatch("SAPI.SpVoice")

browser = webdriver.Chrome("G:/Python_Prog_Examples/chromedriver.exe")
result_val=''
voice_command=''

##Hitting voice to text conversion url
browser.get("https://dictation.io/speech")
time.sleep(12)
clear_text=browser.find_element_by_xpath('//a[@class="btn btn--sm btn-clear btn--primary" and @data-tooltip="Clear Dictation Notepad"]//span[@class="btn__text"]')
#clear the editor window
clear_text.click()
#system speech
speak.Speak("Please click on 'Allow' button and How may I help you..Krishna: ")
start_mic=browser.find_element_by_xpath('//a[@class="btn-mic btn btn--primary-1"]//span[@class="btn__text listen"]')
#click on start mic button
start_mic.click()
#Capturing speech upto 10 seconds
time.sleep(10)
#clicking stop mic button
stop_mic=browser.find_element_by_xpath('//a[@class="btn-mic btn bg--pinterest"]')
stop_mic.click()

#Capturing all voice comands into text format
elem=browser.find_element_by_xpath('//div[@class="ql-editor"]//p')
voice_command=elem.text

speak.Speak("Please wait...I am working on it...")
# aa="please login to facebook" #testing
#Condition for Youtube search
if "YOUTUBE" in voice_command.upper():
    browser.get('https://www.youtube.com/?gl=IN')
    # time.sleep(12)
    browser.implicitly_wait(12)
    you_search = browser.find_element_by_xpath('//input[@id="search"]')
    you_search.send_keys(voice_command)
    you_search.send_keys(Keys.RETURN)
    browser.implicitly_wait(5)
    try:
        video = browser.find_element_by_xpath('//div[@id="title-wrapper"]')
        video.click()
    except:
        #If not got videos link.
        video = browser.find_elements_by_xpath('//div[@id="title-wrapper"]')[1]
        video.click()

#Condition for Facebook search
elif "FACEBOOK" in voice_command.upper():
    browser.get('https://www.facebook.com/')
    browser.implicitly_wait(12)
    uname=''
    pwd=''
    fb_uname_ibox = browser.find_element_by_xpath('//input[@name="email"]')
    fb_uname_ibox.send_keys(uname)
    fb_pwd_ibox = browser.find_element_by_xpath('//input[@name="pass"]')
    fb_pwd_ibox.send_keys(pwd)

    browser.implicitly_wait(5)
    login_btn = browser.find_element_by_xpath('//input[@value="Log In"]')
    login_btn.click()

#Condition for other then youtube search like Google
else:
    browser.get('http://google.com')
    # Hitting the text box
    search = browser.find_element_by_xpath('//input[@class="gLFyf gsfi"]')
    search.send_keys(voice_command)
    search.send_keys(Keys.RETURN)

    try:
        # For any person desription Right Corner
        result_val = browser.find_element_by_xpath('//div[@data-md="50" and @class="mod"]//div/span')
        if result_val != None:
            print(result_val.text)
            speak.Speak(result_val.text)
    except:
        try:
            #For mathemetical calculation
            result_val = browser.find_element_by_xpath('//span[@class="cwcot gsrt"]')
            if result_val != None:
                print(result_val.text)
                speak.Speak(result_val.text)
        except:
            try:
                #For search result box
                result_val=browser.find_element_by_xpath('//span[@class="ILfuVd"]')
                if result_val != None:
                    print(result_val.text)
                    speak.Speak(result_val.text)
            except:
                try:
                    #For Google Map location
                    img_ele = browser.find_element_by_xpath('//div[@class ="lu_map_section"]')
                    if img_ele != None:
                        img_ele.click()
                        time.sleep(2)
                        result_val= browser.find_element_by_xpath('//span[@class="section-facts-description-text"]')
                        if result_val != None:
                            print(result_val.text)
                            speak.Speak(result_val.text)
                except:
                    try:
                        #For website Url's description
                        result_val = browser.find_element_by_xpath('//span[@class="st"]')
                        if result_val != None:
                            print(result_val.text)
                            speak.Speak(result_val.text)
                    except:
                        print("krishna")


Search_result=result_val.text
print("Asked Question is: \n{} \nIt's Answer is: {} : ".format(voice_command,result_val.text))
f_write="Asked Question is:\n" + voice_command + "\nIt's Answer is:\n" + Search_result
#Writting Transcripted text into file.
f= open("Alexa_Conversation.txt","w+")
f.write(f_write)

