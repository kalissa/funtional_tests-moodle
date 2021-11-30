import sys, os, shutil, json, logging, time, csv, re
import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path
from zipfile import ZipFile
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException


sys.path.append(r'source')
from source.dictAttr import DictAttr
from source.azureDataLakeStorageGen1 import AzureDataLakeStorageGen1

class Report(object):
    def __init__(self, org):
        self.org = DictAttr(org)
        self.tempFolder = Path(os.getcwd(), "temp")
        self.baseFolder = Path(os.getcwd(), "bases")
        self.today = datetime.today()
        self.azd = AzureDataLakeStorageGen1(self.org)

    def logDate(self):
        try:
            with open(r'logs\nps_date.txt', 'r') as dt:
                run = dt.readlines()[0]
                self.run = datetime.strptime(run, '%Y-%m-%d')
        except FileNotFoundError:
            sys.exit("Arquivo 'logs\nps_date.txt' não encontrado")

    def excel2csv(self, i,o):
        e = pd.read_excel(i, dtype=str)
        e = e[(e['Date Submitted'] >= str(self.run)) & (e['Date Submitted'] < str(self.today))]
        e.to_csv(o, index=False, quotechar='"', quoting=csv.QUOTE_NONNUMERIC)

    def removeFolder(self):
        if self.tempFolder.exists() and self.tempFolder.is_dir():
            shutil.rmtree(self.tempFolder)

    def uploadDados(self, link):
        #self.logger.info('\n#### Subindo arquivos pro DataLake ####')
        self.azd.connect()
        self.azd.write_azdFile(f'{self.org.adl_nps_folder}\{link}', link)

    def extractZip(self):
        surveyZipFile = max([self.org.DownloadPath + "\\" + f for f in os.listdir(self.org.DownloadPath) if 'survey' in f],key=os.path.getctime)
        with ZipFile(surveyZipFile, 'r') as zf:
            zf.printdir()
            zf.extractall(str(self.tempFolder))
        time.sleep(3) 
        os.remove(surveyZipFile)

    def extract(self):
        driver = webdriver.Chrome(self.org.webDriver)
        driver.get(self.org.page)
        tr = 13 - int((self.today - timedelta(days=1)).month) + 2
        print(tr)

        inputElement = driver.find_element_by_id("email")
        inputElement.send_keys(self.org.hjUser)
        inputElementPass = driver.find_element_by_id("password")
        inputElementPass.send_keys(self.org.hjPsw)
        inputElement.send_keys(Keys.ENTER)
        time.sleep(5)
        
        btn = driver.find_elements_by_xpath('//*[@id="sticky-top-container"]/div[1]/react-primary-nav/nav/div[1]/div[1]')[0].click()
        time.sleep(5)
        role = driver.find_elements_by_xpath('//*[@id="tippy-3"]/div/div/div/div/div/div[2]/ul/li/ul/li[1]')[0].click()
        time.sleep(5)
        try:
            survey = driver.find_elements_by_xpath('//*[@id="site-container"]/react-sidebar/nav/ul/li[5]')[0].click()
            time.sleep(5)

            # if tr == 10: tr = 1  
            
            viewResponses = driver.find_elements_by_xpath(f'//*[@id="content-inner"]/div/div/div[1]/react-surveys-list/div/table/tbody/tr[{tr}]/td[7]/div/a')[0].click()
            time.sleep(5)
            btnCutomize = driver.find_elements_by_xpath(f'//*[@id="content-inner"]/div[3]/div[1]/div/div/react-columns-picker')[0].click()
            time.sleep(5)
            clickHash = driver.find_elements_by_xpath(f'//*[@id="tippy-122"]/div/div/div/div/label[1]/span[1]/span')[0].click()
            time.sleep(1)
            clickBrowser = driver.find_elements_by_xpath(f'//*[@id="tippy-122"]/div/div/div/div/label[6]/span[1]/span')[0].click()
            time.sleep(1)
            clickOS = driver.find_elements_by_xpath(f'//*[@id="tippy-122"]/div/div/div/div/label[7]/span[1]/span')[0].click()
            time.sleep(5)
            iconMoredownloadResponsesXLSX = driver.find_elements_by_xpath('//*[@id="content-inner"]/div[3]/div[1]/div/div/react-export-results')[0].click()
            time.sleep(5)
            downloadResponsesXLSX = driver.find_elements_by_xpath(f'//*[@id="tippy-127"]/div/div/div/ul/li[1]')[0].click()
            time.sleep(35)
            driver.quit()
        except (NoSuchElementException,ElementNotInteractableException) as ex:
            print(ex)
            time.sleep(10)
            self.extract()

    def log_cfg(self):
        pass

    def removeCharacters(self, df, label, regex='\W+|_+'):
        lst = []
        for i in df[label]:
            try:
                rex = [re.sub(regex, '', k).lower() for k in i.split(' ')]
                temp = ' '.join(rex)
            except AttributeError:
                temp = ''
            lst.append(temp)
        df[label] = lst
        return df

    def updateCSV(self, df1, df2, export = True):
        result = pd.concat([df1, df2])
        result = self.removeCharacters(result, 'Como podemos melhorar a sua experiência em nosso Site?')
        if export:
            result.to_csv(f'base_nps.csv', index=False, quotechar='"', quoting=csv.QUOTE_NONNUMERIC)
        return result

    def editCSV(self):
        df = pd.read_csv(f'{self.baseFolder}\\nps_done.csv', dtype=str)
        df['Cluster Area'] = ''

        df=df[['Assunto','Browser','Cluster Area','Como podemos melhorar a sua experiência em nosso Site?',
               'Date Submitted','Device','Em uma escala de 0 a 10, o quanto você recomendaria o Site do Santander para um amigo?',
               'OS','Produto / Servico','QTDE RESP.','Sentimento / Impasse (Nível 1)','Sentimento / Impasse (Nível 2)','Subtema',
               'Ótimo ou Ruim','Source URL']]

        df = df.rename(columns={'Date Submitted':'Data'})
        return df

    def main(self):
        self.removeFolder()
        os.mkdir(self.tempFolder)

        self.logDate()
        self.extract()
        self.extractZip()

        for _, _, files in os.walk(self.tempFolder):
            for f in files:    
                if f.endswith('.xlsx'):
                    excelFile = str(self.tempFolder) + '\\' + f
        csvFile = str(self.baseFolder) + '\\' + 'poll_Date_filtered.csv'
        self.excel2csv(excelFile,csvFile)

        os.system(r'Rscript.exe  --encoding=UTF-8 --no-save .\source\financeira.R')
        os.remove(csvFile)

        df = self.editCSV()
        #base_nps = pd.read_csv('base_nps.csv', dtype=str)
        print('#####\nRemovendo caracteres especiais\n#####')
        result = self.removeCharacters(df, 'Como podemos melhorar a sua experiência em nosso Site?')
        result.to_csv(f'base_nps.csv', index=False, quotechar='"', quoting=csv.QUOTE_NONNUMERIC)
        print('#####\n"base_nps.csv" criado\n#####')
        # _ = self.updateCSV(base_nps, df)
        
        print('#####\nAtualizando arquivo no data lake\n#####')
        s = time.time()
        self.uploadDados("base_nps.csv")
        e = time.time() - s
        print(f'#####\ndone\ntempo de execução: {e}#####')
        os.remove(str(self.baseFolder) + '\\' + 'nps_done.csv')
        os.remove("base_nps.csv")
        self.removeFolder()

if __name__ == '__main__':
    with open('creds\\org.json') as json_file:
        org = json.load(json_file)
    try:
        Report(org).main()
    except FileNotFoundError:
        time.sleep(15)
        Report(org).main()

    run = (datetime.today() - timedelta(days=1)).strftime('%Y-%m-%d')
    with open('logs\\nps_date.txt', 'w') as dt:
        dt.write(run)
