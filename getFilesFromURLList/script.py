import requests
import pandas as pd 
import time

def get_name(url):
    #URL input example: https://cdn.domainon.com/records/2581481-3790983/b6d23bjd-j343-sk22-9987-123kdj2k2jsd-30-11-22-09h36-34612345678-34687654321.mp3
    from_phone = (f'+{url[(url.rfind("/")+53):(url.rfind("-"))]}')
    year = (f'+{url[(url.rfind("/")+44):(url.rfind("/")+46)]}')
    month = (f'+{url[(url.rfind("/")+41):(url.rfind("/")+43)]}')
    day = (f'+{url[(url.rfind("/")+38):(url.rfind("/")+40)]}')
    hour = (f'+{url[(url.rfind("/")+47):(url.rfind("/")+52)]}')
    id = (url[(url.rfind("/")+25):(url.rfind("/")+37)])
    name = f'{from_phone}-{year}{month}{day}-{hour}-{id}.mp3'
    return name #Output name example: +34612345678-221130-09h36-123kdj2k2jsd.mp3

def save_to_csv(file_path, list):
    #Read csv file
    df = pd.read_csv(file_path, header=None)
    #Convert new url to DataFrame
    new_urls_df = pd.DataFrame(list)
    #Add new url to existing DataFrame
    df = pd.concat([df, new_urls_df], ignore_index=True)
    #Save updated DataFarme to csv
    df.to_csv(file_path, index=False, header=False)

def downloadMP3(url):
    response = requests.get(url)

    if response.status_code == 200:
        name = get_name(url)
        #Open file in write mode
        with open(name, 'wb') as file:
            file.write(response.content)
        save_to_csv('./sucess_list.csv', [url])
    else:
        save_to_csv('./error_list.csv', [url])

def download_list_MP3(url_list):
    while url_list:
        n = 0
        for i in range (10):
            if not url_list:
                break
            downloadMP3(url_list.pop(0))
        if not url_list:
            break
        print("sleep 300. Processed lines:")
        n+=10
        print(n)
        time.sleep(300) #Cada 5 MIN

#Open excel file with data (INTRODUCE HERE THE PATH)
f_path = '.list.xlsx' 

#Read data of column A and create a list
df = pd.read_excel(f_path, header=None, usecols="A", engine='openpyxl')
url_list = df[0].dropna().astype(str).tolist()

#Download mp3 files from the url list
download_list_MP3(url_list)