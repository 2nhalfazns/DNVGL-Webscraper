#Created by: Juan R. Cordova, Richard Wu, Sean Park

from selenium import webdriver
import time
 
from tkinter import *

class takeInput(object):

    def __init__(self,requestMessage):
        self.root = Tk()
        self.string = ''
        self.frame = Frame(self.root)
        self.frame.pack()        
        self.acceptInput(requestMessage)

    def acceptInput(self,requestMessage):
        r = self.frame

        k = Label(r,text=requestMessage)
        k.pack(side='left')
        self.e = Entry(r,text='Name')
        self.e.pack(side='left')
        self.e.focus_set()
        b = Button(r,text='Enter',command=self.gettext)
        b.pack(side='right')

    def gettext(self):
        self.string = self.e.get()
        self.root.destroy()

    def getString(self):
        return self.string

    def waitForInput(self):
        self.root.mainloop()

def getText(requestMessage):
    msgBox = takeInput(requestMessage)
    #loop until the user makes a decision and the window is destroyed
    msgBox.waitForInput()
    return msgBox.getString()
 # Asks for what type of ship category you'd like to search for
mySearch = getText('What would you like to search for in the DNV GL Vessel Register?')
print ("Selected Vessel Type: ", mySearch)

newDir = getText('Specify folder Location')
print("Selected directory: ", newDir)

foldName = getText('What is selected folder''s name that you want to save the excel spreadsheet?')
print("Selected Folder: ", foldName)

fileName = getText('What would you like your file to be named?')
print("File Name: ", fileName)
 

driver = webdriver.Firefox() #Utilizes Mozilla Firefox
driver.get('http://vesselregister.dnvgl.com/vesselregister/vesselregister.html')
driver.maximize_window() #Open the driver in a new window, and maximize the window size

time.sleep(.5)

#Clicks on the Search button
driver.find_element_by_class_name('vr-input').send_keys(mySearch)
driver.find_element_by_id('searchBtn').click()
time.sleep(2) #Set to 2s to allow page to load

#Searches for a hyperlink class pertaining to anything that has "vesselid"
a = driver.find_elements_by_xpath('//a[contains(@href,"vesselid")]')
b = [e.get_attribute('href') for e in a]
flags=[True]
for i in range(1,len(b)):
    if b[i-1]==b[i]:
        flags.append(False)
    else:
        flags.append(True)

b = list(a)

#Builds the header, or column titles
titleRow = ('Vessel Name', 'LOA [m]', 'LBP [m]', 'B [m]', 'D [m]', 'T [m]', 'DWT [Tons]')
print(titleRow) #Export Titlerow into Console for easy reading
#Initialize the table which will then be converted into an Excel file
table = []

#b = np.array(b) # Sets the listed hyperlinks found from a to become an array

#Takes the alternating hyperlinks from the Vessel to the next vessel, skipping the "Yes" hyperlink
for i in range(0,len(b)):
    if flags[i]==False:
        continue
    b[i].click()
    time.sleep(2)  #Set to 2s to allow page to load
   
   #Switch to newly opened tab
    curWindowHndl = driver.current_window_handle
    driver.switch_to_window(driver.window_handles[1])
   
   
   
    time.sleep(4)
   
   #Scrape the vessel's name
    vesselName = driver.find_element_by_xpath("//div[contains(@data-bind,'text: name')]").text
   
    time.sleep(1)
  
   #Scroll down to bottom of specific vessel's page
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(3)
  
   #Click on Dimension Tab to open
    dimension = driver.find_element_by_xpath('//a[contains(@href,"#dimensionCollapse")]').click()
    time.sleep(.5)
   
   #Get dimensions
    loa = driver.find_element_by_xpath("//div[contains(@data-bind,'sufixString: loa')]").text
    time.sleep(.1)
   
    lbp = driver.find_element_by_xpath("//div[contains(@data-bind,'sufixString: lbp')]").text
    time.sleep(.1)
   
    B = driver.find_element_by_xpath("//div[contains(@data-bind,'sufixString: b')]").text
    time.sleep(.1)
   
    D = driver.find_element_by_xpath("//div[contains(@data-bind,'sufixString: d')]").text
    time.sleep(.1)
   
    T = driver.find_element_by_xpath("//div[contains(@data-bind,'sufixString: draght')]").text
    time.sleep(.1)
   
    DWT = driver.find_element_by_xpath("//div[contains(@data-bind,'tonString: dwt')]").text
    time.sleep(.1)
   
   #Initiates in defining data as an empty list
    data = []
    data = [vesselName, loa, lbp, B, D, T, DWT]
   
   #Replaces the string ' m' with nothing to get rid of the meter units provided in loa, lbp, B, D, and T values
    for r in range(1,6):
        if data[r].endswith(' m'): #if the measurement scraped from the website contains meter unit, it removes it, and converts the data scraped as an integer
            data[r] =  float(data[r].replace(' m', '')) 
           
   #Replaces the string '' with 'No SamepleData' provided in loa, lbp, B, D, and T values        
    for na in range(0,7):
        if data[na] == '': #if the measurement scraped from the website contains nothing, it changes it to 'No Sample Data'
            data[na] = data[na].replace('', 'No Sample Data')
        else:
            continue
    print(data)
   
   #Export print into list row
    table.insert(i+1,data)

   #Close current Tab  
    driver.close()
   
   #Switch to original tab 
    driver.switch_to_window(curWindowHndl)
    time.sleep(.5)
   
   #Scroll down viewport unit by 29.8 
    driver.execute_script("window.scrollBy(0, 29.8);")
    time.sleep(.1)
    

#Close window   
driver.close()


import pandas
from pandas import DataFrame

#Takes info from table, and sets it as a dataframe
df = pandas.DataFrame.from_records(table, columns = titleRow, coerce_float = True)

#Saved in a given folder within respective laptop
from pandas import ExcelWriter
out_path = newDir + "\\" + foldName + "\\" + fileName + "." + "xlsx"
writer = pandas.ExcelWriter(out_path, engine = 'xlsxwriter')


df.to_excel(writer,'Sheet 1') #df.to_excel
writer.save() #Save file to out_path directory 







