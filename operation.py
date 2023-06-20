import pandas as pd
import numpy as np
import os
import warnings
import time
from selenium.webdriver.common.by import By
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
import requests
from selenium.common.exceptions import NoSuchElementException
from tkinter import *
warnings.filterwarnings('ignore', module="selenium")
class ra_project:
    def __init__ (self):
        self.begin = 0
        self.end = 0
        
    def get_time_range(self):
        def initialize_info():
            self.begin = int(txtfld1.get())
            self.end = int(txtfld2.get())
            window.destroy()
        window=Tk()
        window.title('Peace & Love')
        window.geometry("500x300+10+10")
        lbl=Label(window, text=f"Work Range is :", fg='black', font=("Helvetica", 12))
        lbl.place(x=150, y=50)
        lbl=Label(window, text=f"from", fg='black', font=("Helvetica", 12))
        lbl.place(x=80, y=90)
        txtfld1=Entry(window, text="1", bd=3)
        txtfld1.place(x=120, y=90)
        lbl=Label(window, text=f"to", fg='black', font=("Helvetica", 12))
        lbl.place(x=180, y=90)
        txtfld2=Entry(window, text="2", bd=3)
        txtfld2.place(x=200, y=90)
        btn=Button(window, text="Submit", fg='black',default='active',justify='center',width=10,command=initialize_info)
        btn.place(x=200, y=150)
        # btn.bind_all('<Return>')
        window.mainloop()
        

    def get_code(self):
        # initialize
        url = "https://github.com/omnikdy/ra_file/raw/main/RecallTime%20line_6_19.xlsx"
        r = requests.get(url)
        f = open("current_file.xlsx","wb")
        f.write(r.content)
        f.close()
        data = pd.read_excel('current_file.xlsx')
        data = data.reset_index(drop=True)
        self.code = data['CAMPNO (recall campain ID)']
        #code = code.iloc[::-1].reset_index(drop=True)
        # code = code.iloc[-200:-100]
        self.code = self.code[self.begin:self.end]
        self.code.reset_index(drop=True,inplace=True)
        
    def get_driver(self):
        options = webdriver.ChromeOptions()
        options.add_argument("--incognito")
        options.add_argument("--log-level=OFF")
        options.add_argument("--start-maximized")
        # *************************notice*************************************
        # **************please specify the path below*************************
        prefs = {'profile.default_content_settings.popups': 0,
                    "profile.default_content_setting_values.automatic_downloads": 1,
                    "directory_upgrade": True,
                    "safebrowsing.enabled": True}
        options.add_experimental_option('prefs', prefs)
        options.add_experimental_option('excludeSwitches', ['enable-automation'])
        self.driver = webdriver.Chrome(chrome_options=options,executable_path=ChromeDriverManager().install())
        
    def main_process(self):
        answer = {}
        num = 0
        self.judge = 0
        for new_code in self.code:
            num += 1
            try:
                # next time, just using index
                self.driver.get(f'https://www.nhtsa.gov/recalls?nhtsaId={new_code}')
                handle = self.driver.current_window_handle
                time.sleep(1)
                self.driver.find_element(By.XPATH,'/html/body/div[1]/div[1]/main/article/div/div[3]/div/div/div[2]/div[1]/section/div[2]/div/div[1]/div/div/div[1]/div/a').click()
                time.sleep(1)
                self.driver.find_element(By.XPATH,'/html/body/div[1]/div[1]/main/article/div/div[3]/div/div/div[2]/div[1]/section/div[2]/div/div[1]/div/div/div[2]/div[1]/div[4]/div[2]/a').click()
                time.sleep(3)
                # key work is Defect and Noncompliance
                # judge which document is needed for us
                document_name = self.driver.find_elements(By.XPATH,'//*[@id="associated-documents-58"]/ul/li')
                index = []
                for i in range(len(document_name)):
                    if 'Defect' in document_name[i].text or 'Noncompliance' in document_name[i].text:
                        index.append(i+1)
                # if index == 0: print("no required files") 
                for i in index: 
                    link = self.driver.find_element(By.XPATH,f'//*[@id="associated-documents-58"]/ul/li[{i}]/a').get_attribute("href")
                    new=f'window.open("{link}");'
                    self.driver.execute_script(new)
                time.sleep(1)
                # answer box
                def extract_info():
                    answer[new_code]=txtfld.get()
                    window.destroy()
                def break_down():
                    answer[new_code]=txtfld.get()
                    window.destroy()
                    self.driver.quit()
                    self.judge = 1
                window=Tk()
                window.title('Peace & Love')
                window.geometry("500x300+10+10")
                lbl=Label(window, text=f"Current code : {new_code}", fg='black', font=("Helvetica", 12))
                lbl.place(x=150, y=50)
                lbl=Label(window, text=f"You have finished : {num}", fg='black', font=("Helvetica", 12))
                lbl.place(x=150, y=70)
                lbl=Label(window, text=f"The answer is :", fg='black', font=("Helvetica", 12))
                lbl.place(x=105, y=100)
                txtfld=Entry(window, text="This is Entry Widget", bd=5)
                txtfld.place(x=230, y=100)
                btn=Button(window, text="Next", fg='black',default='active',justify='center',width=10,command=extract_info)
                btn.place(x=200, y=150)
                btn=Button(window, text="Quit", fg='black',default='active',justify='center',width=10,command=break_down)
                btn.place(x=200, y=200)
                window.mainloop()
                if self.judge==1:
                    break
                handles = self.driver.window_handles
                for newhandle in handles[1:]:
                    self.driver.switch_to.window(newhandle)
                    self.driver.close()
                self.driver.switch_to.window(handle)
            except NoSuchElementException:
                answer[new_code]='missing file'
        self.driver.quit()
        index = 0
        while os.path.exists(f'{time.strftime("%Y_%m_%d", time.localtime())}_{index}.npy'):
            index += 1
        np.save(f'{time.strftime("%Y_%m_%d", time.localtime())}_{index}.npy', answer)
        
if __name__ == "__main__":
    dealer = ra_project()
    dealer.get_time_range()
    dealer.get_code()
    dealer.get_driver()
    dealer.main_process()