##BY hongvin
##github.com/khvmaths
COURSE_CODE = 'KIE4014'
REFRESH_TIME = 3

from bs4 import BeautifulSoup
import pandas as pd
import pyttsx3
import requests
import time
import datetime
import threading

site = "http://register.um.edu.my/jadual_penawaranbi.asp?s_Sesi=2019%2F2020&s_Semester=1&s_fakulti=0&s_SLOT_KOD_SUBJEK="+COURSE_CODE
engine = pyttsx3.init()

def startEngine():
    t=threading.Timer(REFRESH_TIME,startEngine).start()

    print('['+str(datetime.datetime.now().time())+']: List refreshed!')
    try:
        response = requests.get(site)
    except:
            print('error!')
    soup = BeautifulSoup(response.text,"html.parser")
    tr = soup.find_all('tr')
    td= tr[8].find_all('td')

    for br in soup.find_all('br'):
        br.replace_with(' ')

    pp=[]
    for i in td[4:7]:
        try:
            pp.append(i.text)
        except:
            print('Error lol')

    group=pp[0].split()
    capacity=pp[1].split()
    vacancy=pp[2].split()

    df=pd.DataFrame({'Capacity':capacity,'Vacancy':vacancy})
    df.index=group
    df['Capacity']=df['Capacity'].astype(int)
    df['Vacancy']=df['Vacancy'].astype(int)
    print(df)

    for i in range(len(group)):
        if not (df.iat[i,1] == 0):
            engine.say('Vacancy found Group'+str(i+1))
            print('['+str(datetime.datetime.now().time())+']: FOUND!: GROUP '+str(i))
            engine.runAndWait() 

startEngine()
