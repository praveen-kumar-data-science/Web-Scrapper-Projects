#!/usr/bin/env python
# coding: utf-8

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
import re
import time
import pandas as pd
from bs4 import BeautifulSoup
from openpyxl import Workbook
import xlsxwriter
import csv
import warnings
warnings.filterwarnings('ignore')
from IPython.display import clear_output

PATH_TO_CHROME_DRIVER = 'chromdriver.exe'

# This files stores the linkedin username and password details, go to the file and update the details of your linkedin credentials
CREDENTIALS_FILE = 'credentials2.txt'
# DOWNLOAD_PATH = ''

# We wanted to extract the data of CEOs of various firms around the world - firm and CEO names are present in this csv file
df = pd.read_csv('firmid_emp_allgroup_scrape_01222023.csv')
df['Parent company'] = df['Parent company'].str.split('(').str[0]
df['contact_person'] = df['contact_person'].str.split('and').str[0]

def append_to_excel(fpath, df, sheet_name):
    with pd.ExcelWriter(fpath, mode="a", engine="openpyxl") as f:
        df.to_excel(f, sheet_name=sheet_name)

# create a workbook as .xlsx file
def create_workbook(path):
    workbook = Workbook()
    workbook.save(path) 

def linkedin_login(EXE_PATH, CREDENTIALS_FILE):
    
    with open(CREDENTIALS_FILE) as f:
        details = f.read()
        
    USERNAME, PASSWORD = details.split('\n')

    driver = webdriver.Chrome(EXE_PATH)
    driver.get("https://www.linkedin.com/uas/login")
    email = driver.find_element(By.ID, "username")
    email.send_keys(USERNAME)
    password=driver.find_element(By.ID, "password")
    password.send_keys(PASSWORD)
    password.send_keys(Keys.RETURN)
    
    return driver
    
def linkedin_scrape(id_, driver, contact_obtained, company_obtained, firm_id):
    # Get the exact linkedin profile link of the person 
    time.sleep(3)
    search_field = driver.find_element("xpath", "//input[@aria-label='Search']")
    search_field.clear()
    time.sleep(2)
    
    query = contact_obtained + ' ' + company_obtained
    
    search_field.send_keys(query)
    search_field.send_keys(Keys.RETURN)
    time.sleep(3)
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='People']"))).click()
    time.sleep(3)
    
    # If there is no linkedin profile for the person, we will skip
    if 'No results found' in str(driver.page_source):
        pass
    else:
        fname = query + '.xlsx'
        EXCEL_PATH = f'/Users/natexu/Documents/Dissertation/Chapter3/Linkedin_data/nate_files/scrape_excels/' + fname 
        create_workbook(path=EXCEL_PATH)
        
        soup = BeautifulSoup(driver.page_source, features='lxml')
        links = soup.findAll("span", {"class": "entity-result__title-line entity-result__title-line--2-lines"})
        for link in links:
            new_link = link.find('a').get('href').split('?')[0]
            break
        driver.get(new_link)
        time.sleep(3)
        
        # About Section
        soup = BeautifulSoup(driver.page_source, features='lxml')
        abouts = soup.findAll("div", {"class": "display-flex ph5 pv3"})
        if abouts != []: 
            for about in abouts:
                about_text = about.text
                break
            about_text = ' '.join(about_text.strip().split('\n'))
        else:
            about_text = 'NONE'

        # Followers
#         return soup
        for k in soup.findAll('p', {'class':"pvs-header__subtitle pvs-header__optional-link text-body-small"}):
            followers = k.find('span', {'class':'visually-hidden'}).text.split('followers')[0].replace(',', '')
        try:
            followers = int(followers)
        except:
            followers = 'NONE'
        #if len(str(followers)) < 1: followers = 'NONE'
        
        # Languages section
        driver.get(new_link + '/details/languages')
        time.sleep(2)
        soup = BeautifulSoup(driver.page_source, features='lxml')
        languages = []
        lang_cnts = soup.findAll('div', {'class': 'scaffold-finite-scroll__content'})
        for lang_cnt in lang_cnts:
            for lang in lang_cnt.findAll('span', {'class':'mr1 t-bold'}):
                languages.append(lang.find('span', {'class':'visually-hidden'}).text)
        languages = ', '.join(languages)       
        if languages.strip() =='': languages = "NONE"

        # Append to excel
        df = pd.DataFrame()
        df['unique_id'] = [id_]
        df['contact_person'] = [contact_obtained]
        df['firm_id'] = firm_id
        df['contact_company'] = [company_obtained]
        df['about'] = [about_text]
        df['followers'] = [followers]
        df['url'] = [new_link]
        df['languages'] = [languages]
        append_to_excel(EXCEL_PATH, df, 'firm_contact')
       
        # Education section
        driver.get(new_link + '/details/education')
        time.sleep(2)
        soup = BeautifulSoup(driver.page_source, features='lxml')
        edu = soup.findAll('div', {'class': 'display-flex flex-column full-width align-self-center'})

        colleges_, all_times_, all_texts_, degs = [], [], [],[]
        for e in edu:
            clgs = e.findAll('span', {'class': 'mr1 hoverable-link-text t-bold'})
            for clg in clgs:
                colleges_.append(clg.find('span', {'class': 'visually-hidden'}).text)
            times = e.findAll('span', {'class': 't-14 t-normal t-black--light'})

            if "t-14 t-normal t-black--light" in str(e):
                for t in times:
                    all_times_.append(t.find('span', {'class': 'visually-hidden'}).text)
            else:
                all_times_.append('NONE')
            degrees = e.findAll('span', {'class':'t-14 t-normal'})
            if 't-14 t-normal' in str(e):
                for deg in degrees:
                    degs.append(deg.find('span', {'class':'visually-hidden'}).text)
            else:
                degs.append('NONE')
            new_str = ''
            texts = e.findAll('div', {'class': "pvs-list__outer-container"})
            if "pv-shared-text-with-see-more full-width t-14 t-normal t-black display-flex align-items-center" in str(e):
                for text in texts:
                    try:
                        new_str = new_str + ' ' + text.find('span', {'class': 'visually-hidden'}).text
                    except:
                        pass
            else:
                new_str = 'NONE'
            all_texts_.append(new_str.strip())
        
        # Append to excel
        df = pd.DataFrame()
        df['Education_entity'] = colleges_
        df['Education_degree'] = degs
        df['Education_time'] = all_times_
        df['contact_person'] = contact_obtained
        df['firm_id'] = firm_id
        df['contact_company'] = company_obtained
        df['unique_id'] = id_
        append_to_excel(EXCEL_PATH, df, 'Education')

        # Certifications section
        driver.get(new_link + '/details/certifications')
        time.sleep(2)
        soup = BeautifulSoup(driver.page_source, features='lxml')
        certifications, cert_entities, issued, new_issued, expired = [],[],[],[],[]
        certificates = soup.findAll('div', 
                                    {'class':"scaffold-layout__inner scaffold-layout-container scaffold-layout-container--reflow"})
        for certificate in certificates:
            if 'mr1 t-bold' in str(certificate):
                for c in certificate.findAll('span', {'class':'mr1 t-bold'}):
                    certifications.append(c.find('span', {'class':'visually-hidden'}).text)
            else:
                for c in certificate.findAll('span', {'class':'mr1 hoverable-link-text t-bold'}):
                    certifications.append(c.find('span', {'class':'visually-hidden'}).text)
            for i, c in enumerate(certificate.findAll('span', {'class':'t-14 t-normal t-black--light'})):
                l = c.find('span', {'class':'visually-hidden'}).text
                if i%2 ==0:
                    issued.append(l)
            for c in certificate.findAll('span', {'class':'t-14 t-normal'}):
                cert_entity = c.find('span', {'class':'visually-hidden'}).text.split(' · ')
                cert_entities.append(cert_entity[0])
        
        for issue in issued:
            if 'Expires' in issue:
                i, j = issue.split('Expires')[0], issue.split('Expires')[1]
            else:
                j ='NONE'
                i = 'NONE' if issue.startswith('Credential') else issue
            new_issued.append(i)
            expired.append(j)
        if len(certifications) > len(new_issued):
            new_issued += ['NONE'] * (len(certifications) - len(new_issued))
            expired += ['NONE'] * (len(certifications) - len(new_issued))
        if len(new_issued) > len(cert_entities): cert_entities += (len(new_issued) - len(cert_entities)) * ['NONE']
        if len(certifications) > len(new_issued): new_issued += (len(certifications) - len(new_issued))*['NONE']
        if len(certifications) > len(expired): expired += (len(certifications) - len(expired))*['NONE']
        if len(certifications) > len(cert_entities): cert_entities += (len(certifications) - len(cert_entities))*['NONE']
        #if len(certifications) > len(cert_entities): cert_entities += (len(certifications) - len(cert_entities))*['NONE']
            
        # Append to excel
        df = pd.DataFrame()
        df['title'] = certifications
        df['cert_entiry'] = cert_entities
        df['issued'] = new_issued
        df['expires'] = expired
        df['contact_person'] = contact_obtained
        df['firm_id'] = firm_id
        df['contact_company'] = company_obtained
        df['unique_id'] = id_
        append_to_excel(EXCEL_PATH, df, 'Licenses&Certifications')
        
        # Volunteering section
        driver.get(new_link + '/details/volunteering-experiences')
        time.sleep(3)
        soup = BeautifulSoup(driver.page_source, features='lxml')
        volunteer_exps, volunteer_roles, volunteer_times = [],[],[]
        voltrs = soup.findAll('div', {'class': 'scaffold-finite-scroll__content'})
        for volunteering in voltrs:
            for volunteer in volunteering.findAll('span', {'class':'mr1 t-bold'}):
                volunteer_exps.append(volunteer.find('span', {'class':'visually-hidden'}).text)
            for volunteer in volunteering.findAll('span', {'class':'t-14 t-normal'}):
                volunteer_roles.append(volunteer.find('span', {'class':'visually-hidden'}).text)
            for i, volunteer in enumerate(volunteering.findAll('span', {'class':'t-14 t-normal t-black--light'})):
                if i%2==0: volunteer_times.append(volunteer.find('span', {'class':'visually-hidden'}).text)

        # Append to excel
        df = pd.DataFrame()
        df['Volunteering_entity'] = volunteer_exps
        df['Volunteering_role'] = volunteer_roles
        df['Volunteering_time'] = volunteer_times
        df['contact_person'] = contact_obtained
        df['firm_id'] = firm_id
        df['contact_company'] = company_obtained
        df['unique_id'] = id_
        append_to_excel(EXCEL_PATH, df, 'volunteer_roles')
        
        # Experience section
        driver.get(new_link + '/details/experience')
        time.sleep(3)
        soup = BeautifulSoup(driver.page_source, features='lxml')
        exps = soup.findAll('div', {'class': 'display-flex flex-column full-width align-self-center'})
        companies, roles, exp_times, exp_locs, descs, temp = [], [], [], [], [],[]
        multiplicate = 0
        for i, exp in enumerate(exps):
            remember = ''
            if "mr1 hoverable-link-text t-bold" in str(exp):
                rols = exp.findAll('span', {'class': 'mr1 hoverable-link-text t-bold'})
                for role in rols:
                    temp.append(role.find('span', {'class': 'visually-hidden'}).text)
                if len(temp) > 1: 
                    remember = temp[0]
                    multiplicate = len(temp) - 1
                roles += temp[1:]
                temp=[]
            elif "mr1 t-bold" in str(exp):
                rols = exp.findAll('span', {'class': 'mr1 t-bold'})
                for role in rols:
                    roles.append(role.find('span', {'class': 'visually-hidden'}).text)
            else:
                roles.append('NONE')

            times = exp.findAll('span', {'class': 't-14 t-normal t-black--light'})

            if "t-14 t-normal t-black--light" in str(exp):
                for i, t in enumerate(times):
                    if i%2 ==0:
                        exp_times.append(t.find('span', {'class': 'visually-hidden'}).text)
                    else:
                        exp_locs.append(t.find('span', {'class': 'visually-hidden'}).text)
            else:
                exp_times.append('NONE')

            if remember!='':
                companies += [remember]  * multiplicate
                remember = ''

            else:
                if "t-14 t-normal" in str(exp):
                    cpys = exp.findAll('span', {'class': "t-14 t-normal"})
                    for cpy in cpys:
                        x = cpy.find('span', {'class': 'visually-hidden'}).text
                        companies.append(x)
                elif "t-14 t-normal t-black--light" in str(exp):
                    cpys = exp.findAll('span', {'class': "t-14 t-normal t-black--light"})
                    for cpy in cpys:
                        companies.append(cpy.find('span', {'class': 'visually-hidden'}).text)

            if 'display-flex align-items-center t-14 t-normal t-black' in str(exp):
                descrips = exp.findAll('div', {'class':'display-flex align-items-center t-14 t-normal t-black'})
                for desc in descrips:
                    descs.append(desc.text.replace('\n', '').strip())
            else:
                descs.append('NONE')
        list_to_remove = ['Full-time', 'Part-time', 'Self-employed', 'Internship', 'Freelance', 'Contract', 'Apprenticeship','Seasonal']
        companies = [company for company in companies if company not in list_to_remove]
        descs_ = []
        uniques = list(dict.fromkeys(descs))
        for desc in descs:
            if desc == 'NONE' or desc in uniques:
                if desc not in descs_:
                    descs_.append(desc)
        if len(roles) > len(descs_):
            diff = len(roles) - len(descs_)
            descs_ += ['NONE'] * diff
        if len(exp_locs) < len(companies): exp_locs += exp_locs + ['NONE']*(len(companies) - len(exp_locs))
        if len(roles) > len(companies): companies += ['NONE']*(len(roles) > len(companies))
        k = min(len(companies), len(roles), len(exp_times), len(exp_locs), len(descs))
        companies, roles, exp_times, exp_locs, descs = companies[:k], roles[:k], exp_times[:k], exp_locs[:k], descs_[:k]
        exp_dur,exp_period = [],[]
        for exp_time in exp_times:
            e = exp_time.split(' · ') 
            if len(e) > 1:
                exp_dur.append(e[1])
                exp_period.append(e[0])
            else:
                exp_dur.append(e[0])
                exp_period.append(e[0])
        
        # Append to excel
        df = pd.DataFrame()
        df['Experience_entity'] = companies
        df['Experience_title'] = roles
        df['Experience_location'] = exp_locs
        df['Experience_time'] = exp_period
        df['Experience_duration'] = exp_dur
        df['content'] = descs
        new_companies = []
        for c in companies:
            c_last = c.split(' · ')[-1]
            if c_last not in list_to_remove:
                new_companies.append('N/A')
            else:
                new_companies.append(c_last)
                
        df['full_time_part_time'] = new_companies
        df['contact_person'] = contact_obtained
        df['firm_id'] = firm_id
        df['contact_company'] = company_obtained
        df['unique_id'] = id_
        append_to_excel(EXCEL_PATH, df, 'Experience')
        
        # Honors and Awards section
        driver.get(new_link + '/details/honors')
        time.sleep(3)
        soup = BeautifulSoup(driver.page_source, features='lxml')

        awards, issuing_entities, issue_times = [],[],[]
        all_awards = soup.findAll('div', {'class': 'scaffold-finite-scroll__content'})
        for award_ in all_awards:
            for award in award_.findAll('span', {'class':'mr1 t-bold'}):
                awards.append(award.find('span', {'class':'visually-hidden'}).text)
            for award in award_.findAll('span', {'class':'t-14 t-normal'}):
                award_txt = award.find('span', {'class':'visually-hidden'}).text.split(' · ')
                issuing_entities.append(award_txt[0])
                try:
                    issue_times.append(award_txt[1])
                except:
                    issue_times.append('NONE')
        
        # Append to excel
        df = pd.DataFrame()
        df['awards_title'] = awards
        df['issue_entity'] = issuing_entities
        df['issue_time'] = issue_times
        df['contact_person'] = contact_obtained
        df['firm_id'] = firm_id
        df['contact_company'] = company_obtained
        df['unique_id'] = id_
        append_to_excel(EXCEL_PATH, df, 'Honors&Awards')
        
        # Organizations section
        driver.get(new_link + '/details/organizations')
        time.sleep(3)
        soup = BeautifulSoup(driver.page_source, features='lxml')

        organizations, org_times, org_roles = [],[],[]
        all_organizations = soup.findAll('div', {'class': 'scaffold-finite-scroll__content'})
        for organization__ in all_organizations:
            for org in organization__.findAll('span', {'class':'mr1 t-bold'}):
                organizations.append(org.find('span', {'class':'visually-hidden'}).text)
            for org in organization__.findAll('span', {'class':'t-14 t-normal'}):
                org_time = org.find('span', {'class':'visually-hidden'}).text.split(' · ')
                if len(org_time)>1:
                    org_time = org_time[1]
                else:
                    org_time = org_time[0]
                org_times.append(org_time)
#                 org_roles.append(org_role)
        org_times = [time_ if bool(re.search(r'\d', time_)) else 'NONE' for time_ in org_times]

        # Append to excel
        df = pd.DataFrame()
        df['Organizations'] = organizations
        df['Organizations_time'] = org_times
        df['contact_person'] = contact_obtained
        df['firm_id'] = firm_id
        df['contact_company'] = company_obtained
        df['unique_id'] = id_
        append_to_excel(EXCEL_PATH, df, 'Organizations')
        
        # Recommendations section
        driver.get(new_link + '/details/recommendations')
        time.sleep(2)
        soup = BeautifulSoup(driver.page_source, features='lxml')

        #### Received
        received_names, received_recommendations, received_relationships, received_times = [],[],[],[]
        all_recs = soup.findAll('div', {'id': 'ember47'})
        for rec_ in all_recs:
            for rec in rec_.findAll('span', {'class':'mr1 hoverable-link-text t-bold'}):
                received_names.append(rec.find('span', {'class':'visually-hidden'}).text)
            for rec in rec_.findAll('div', {'class', 'display-flex align-items-center t-14 t-normal t-black'}):
                received_recommendations.append(rec.text)
            for rec in rec_.findAll('span', {'class', 't-14 t-normal t-black--light'}):
                t = rec.find('span', {'class':'visually-hidden'}).text.split(',')
                received_relationships.append(t[-1].strip())
                received_times.append(' '.join(t[:-1]))
        #### Given
        given_names, given_recommendations, given_relationships, given_times = [],[],[],[]
        all_recs = soup.findAll('div', {'id': 'ember45'})
        for rec_ in all_recs:
            for rec in rec_.findAll('span', {'class':'mr1 hoverable-link-text t-bold'}):
                given_names.append(rec.find('span', {'class':'visually-hidden'}).text)

            for rec in rec_.findAll('div', {'class', 'display-flex align-items-center t-14 t-normal t-black'}):
                given_recommendations.append(rec.text)
            for rec in rec_.findAll('span', {'class', 't-14 t-normal t-black--light'}):
                t = rec.find('span', {'class':'visually-hidden'}).text.split(',')
                given_relationships.append(t[-1].strip())
                given_times.append(' '.join(t[:-1]))
                
        # Append to excel
        received_names, received_recommendations, received_relationships, received_times
        df = pd.DataFrame()
        df['recommender'] = received_names
        df['recommendation'] = received_recommendations
        df['relationship'] = received_relationships
        df['time'] = received_times
        df['contact_person'] = contact_obtained
        df['firm_id'] = firm_id
        df['contact_company'] = company_obtained
        df['unique_id'] = id_
        append_to_excel(EXCEL_PATH, df, 'Recommendation_received')        

        df = pd.DataFrame()
        df['recommended'] = given_names
        df['recommendation'] = given_recommendations
        df['relationship'] = given_relationships
        df['time'] = given_times
        df['contact_person'] = contact_obtained
        df['firm_id'] = firm_id
        df['contact_company'] = company_obtained
        df['unique_id'] = id_
        append_to_excel(EXCEL_PATH, df, 'Recommendation_given')   
        
        # Skills section
        driver.get(new_link + '/details/skills')
        time.sleep(3)
        soup = BeautifulSoup(driver.page_source, features='lxml')

        skills, skills_endorsements = [],[]
        all_skills = soup.findAll('ul', {'class':'pvs-list'})
        for s in all_skills:
            for skill in s.findAll('span', {'class':'mr1 hoverable-link-text t-bold'}):
                skills.append(skill.find('span', {'class':'visually-hidden'}).text)

            for skill in s.findAll('div', {'class':'display-flex'}):
                skills_endorsements.append(skill.text.strip().replace('\n',''))

        endorsements = []
        for skill in skills:
            endorsement = ''
            for skill_endorsement in skills_endorsements:
                if skill in skill_endorsement:
                    if 'endorsement' in skill_endorsement:
                        
                        endorsement = int(re.findall("([0-9]*) endorsement", skill_endorsement)[0])
                        endorsements.append(endorsement)
                        break
            if endorsement == '':
                endorsements.append('NONE')
                
        df = pd.DataFrame()
        df['skills'] = skills
        df['skill_endorsements'] = endorsements
        df['unique_id'] = id_
        df['contact_person'] = contact_obtained
        df['firm_id'] = firm_id
        df['contact_company'] = company_obtained
        append_to_excel(EXCEL_PATH, df, 'Skills')   

# RUN THIS FIRST IF THERE IS A BOT\n",
# IF NO BOT JUST RUN THIS CELL AND THE NEXT CELL INSTANTLY! NO NEED TO WAIT..\n",

# MAKE SURE THE DRIVER CHROME PAGE (TARGET WINDOW) IS OPEN ALL TIME OR ELSE RUN THIS CELL AGAIN AND RUN NEXT CELL..\n",

def where_to_start(file):
    with open(file) as f:
        return int(f.read())
START = where_to_start('profile_count.txt')
END = len(df)
driver = linkedin_login(PATH_TO_CHROME_DRIVER, CREDENTIALS_FILE)
time.sleep(2)

# WAIT FOR THE PREVIOUS CELL AND THEN RUN THIS CELL

for row in df[START:END].iterrows():
    index = row[0]
    firm_id = str(row[1]['firm_ID']).strip()
    contact, company = row[1]['contact_person'].strip(), row[1]['Parent company'].strip()
    
    if index%5 == 0:
        driver = linkedin_login(PATH_TO_CHROME_DRIVER, CREDENTIALS_FILE)
        clear_output(wait=True)
    try:
        linkedin_scrape(index, driver, contact, company, firm_id)
        time.sleep(2)
    except Exception as e:
        print(e)
        # If there is any error in the profile scrapping, then add those profile index into a separate csv file 'skipped_profiles.csv'
        with open('skipped_profiles.csv', 'w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow([index, contact, company, firm_id])
        pass
    with open('profile_count.txt', 'w') as f:
        #f.truncate()
        print('Last index for START is: ', index)
        f.write(str(index))
        f.close()
