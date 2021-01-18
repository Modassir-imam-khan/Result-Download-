from selenium import webdriver
from helium import *
import pandas as pd
from pathlib import Path
import os
import time
import numpy as np
from collections import OrderedDict
from functools import partial
import pdb
import shutil

from utils import *

captcha_loaded = partial(img_loaded, xpath='//*[@id="ctl00_ContentPlaceHolder1_captchaimg"]')

def wait_until_img_loads(browser, xpath):
    check_loaded = partial(img_loaded, xpath=xpath)
    loaded = check_loaded(browser)
    while not loaded:
        time.sleep(2)
        loaded = check_loaded(browser)

def open_page():
    browser = start_chrome()
    browser.get("http://ntaresults.nic.in/NTARESULTS_CMS/public/home.aspx")
    click(Text('JEE(Main)-2020 NTA Score-Paper 1 (B.E./B.Tech.)'))
    window_after = browser.window_handles[1]
    browser.switch_to.window(window_after)
    browser.refresh()
    
    loaded = captcha_loaded(browser)
    while not loaded:
        browser.implicitly_wait(200)
        browser.refresh()
        browser.implicitly_wait(200)
        loaded = captcha_loaded(browser)
        
    browser.implicitly_wait(200)
    return browser
 

def fill_details(app_no, dob="day-month-year", browser=None):
    day, month, year = dob.split("-")
    app_no_field = browser.find_elements_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_txtRegNo"]')[0]
    app_no_field.send_keys(app_no)
    day_menu = browser.find_elements_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_ddlday"]')[0]
    day_menu.send_keys(day)
    
    month_options = {"01": "January (01)", "02": "Februrary (02)", "03": "March (03)", "04": "April (04)", 
                     "05": "May (05)", "06": "June (06)", "07": "July (07)", "08": "August (08)", 
                     "09": "September (09)", "10": "October (10)", "11": "November (11)", "12": "December (12)"}
    
    month_menu = browser.find_elements_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_ddlmonth"]')[0]
    month_menu.send_keys(month_options[month])
    year_menu = browser.find_elements_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_ddlyear"]')[0]
    year_menu.send_keys(year)
    return browser


def extract_info(browser):
    data = extract_table(browser, 2)
    tables = pd.read_html(browser.page_source)

    scores = tables[4].iloc[2:, 1:-1]
    scores.columns = tables[4].iloc[1, 1:-1]
    scores.index = ["Jan", "Sep", "Final"] #list(tables[4].iloc[2:, 0])
    for sess in scores.index:
        for col in list(scores.columns):
            data[f'{col}-{sess}'] = scores.loc[sess, col]
    

    ranks = tables[6].iloc[2:, 1:]
    ranks.columns = tables[6].iloc[1, 1:]
    ranks.index = [0]
    for cat in ranks.columns:
        data[f"{cat}-rank"] = ranks.loc[0, cat]

    return data


def extract_table(browser, table, data=None):
    if data is None: data = OrderedDict()
    table_xpath = f'//*[@id="tableToPrint"]/tbody/tr/td/table[{table}]/tbody'
    row_count = len(browser.find_elements_by_xpath(f'{table_xpath}/tr'))
    for i in range(1, row_count+1):
        col_count = len(browser.find_elements_by_xpath(f'{table_xpath}/tr[{i}]/td'))
        for j in range(1, col_count, 2):
            k = browser.find_element_by_xpath(f'{table_xpath}/tr[{i}]/td[{j}]').text
            v = browser.find_element_by_xpath(f'{table_xpath}/tr[{i}]/td[{j+1}]').text
            data[k] = v
    return data


def get_data_mains(app_no, dob, *args, **kwargs):
    browser = open_page()
    browser.implicitly_wait(10)
    browser = fill_details(app_no, dob, browser)
    wait_until(Text('State Code of Eligibility').exists, timeout_secs=100)
    data = extract_info(browser)
    return data, browser


def save_screenshot(browser, img_file, fullpage=True): 
    if fullpage:
        # from Screenshot import Screenshot_Clipping
        # browser.maximize_window()
        # ob=Screenshot_Clipping.Screenshot()
        # ob.full_Screenshot(browser, save_path=str(img_file.parent), image_name=img_file.name, load_wait_time=1, )
        fullpage_screenshot(browser, str(img_file))
        # browser.find_element_by_tag_name('body').screenshot(str(img_file))   
    else:
        browser.set_window_size(1920, 1280)
        browser.save_screenshot(str(img_file)) 


def get_data_advanced(app_no, dob, phone, *args, **kwargs):
    browser = start_chrome()
    go_to('http://result.jeeadv.ac.in/index.php')

    roll_no_field = S('//*[@id="results-form"]/table/tbody/tr[1]/td[2]/input').web_element
    roll_no_field.clear()
    roll_no_field.send_keys(app_no)

    dob_field = S('//*[@id="results-form"]/table/tbody/tr[2]/td[2]/input').web_element
    dob_field.send_keys(dob.replace('/', '-'))

    phone_field = S('//*[@id="results-form"]/table/tbody/tr[3]/td[2]/input').web_element
    phone_field.send_keys(phone)

    wait_until(Text('Positive Score').exists, timeout_secs=500)

    dfs = pd.read_html(browser.page_source)
    if not Text('You have not qualified JEE (Advanced) 2020.').exists():
        df = pd.concat([dfs[1].transpose().rename(columns=dfs[1].transpose().iloc[0]).drop(dfs[1].index[0]).reset_index(drop=True).transpose(), dfs[2].transpose()])
    else:
        df = dfs[1].transpose()
    data = df.to_dict()[0]

    return data, browser


def get_data_neet(app_no, dob, phone, roll_no, *args, **kwargs):
    browser = start_chrome('http://ntaresults.nic.in/NEET20/Result/ResultNEET.htm')
    write(str(roll_no).strip(), into='Enter Roll Number')
    write(str(dob).replace('-', '/'), into='Enter DOB(dd/mm/yyyy)')
    wait_until(Text('Senior Director [NEET (UG)]').exists, timeout_secs=100)
    dfs = pd.read_html(browser.page_source)
    
    data = OrderedDict()
    df = dfs[5]
    df = df.dropna(how='all', axis=1)

    data.update(df[[0, 1]].set_index(0).to_dict()[1])
    data.update(df[[2, 3]].set_index(2).to_dict()[3])
    
    df = dfs[7]
    df = df.dropna(how='all', axis=1)

    data.update(df[[0, 1]].drop(0).set_index(0).to_dict()[1])
    data.update(df[[2, 3]].drop([1, 3, 4]).reset_index(drop=True).set_index(2).to_dict()[3])
    
    data.update(dfs[10].transpose().set_index(0).to_dict()[1])
    return data, browser


def get_neet_admit_data(app_no, dob, *args, **kwargs):
    browser = start_chrome('https://ntaneet.nic.in/ntaneet/admitcard/admitcard.html')
    write(str(app_no).strip(), into='Enter Application Number')
    write(str(dob).replace('-', '/'), into='Enter DOB(dd/mm/yyyy)')
    t = "The above undertaking has to be filled in advance before reaching the centre, except candidate signature which has to be done in the presence of Invigilator."
    wait_until(Text(t).exists, timeout_secs=500, interval_secs=1)
    wait_until_img_loads(browser, '//*[@id="imgCandPhoto"]')

    dfs = pd.read_html(browser.page_source)
    data = {k.strip(':').replace("’","'"): v for k, v in dfs[5].dropna(how='all', axis=1)[[0, 1]].iloc[[0, 1]].set_index(0).to_dict()[1].items()}
    # click(S('//*[@id="tblAppno"]/tbody/tr[4]/td[2]/img[2]'))
    data = {"Roll No.": data['Roll Number'], "Candidate's Name": data["Candidate's Name"]}
    return data, browser



def neet_conf_info(a, b, i, data={}):
    if a != b and str(a) != "nan" and str(b) != "nan" and (a not in ['Pnoto']):
        if i == 14 or (a.strip().startswith("(") and a.strip().endswith(')') and b.strip().startswith("(") and b.strip().endswith(')')):
            if "Choice of Examination City" not in data.keys():
                data['Choice of Examination City'] = ""
            data['Choice of Examination City'] = data['Choice of Examination City'] + f"{a};{b};"
        elif a.startswith('Transaction ID.') and b.startswith('Total Amount'):
            data['Transaction ID.'] = a.split(":")[1].strip()
            data['Total Amount'] = b.split(":")[1].strip()
        elif a.startswith("Application No:"):
            data['Application No.'] = b.strip()
        else:
            if "(" in a and ")" in a:
                a = a[a.index('(')+1: a.index(')')]
            if a == 'Date of Birth':
                b = b.strip().replace('/', '-')
            data[a.strip()] = b.strip()
    return data


def login_neet_conf_website(df, i):
    app_no = get_prop(df, i, "Application No.")
    passwd = get_prop(df, i, "Password")
    browser = start_chrome('https://ntaneet.nic.in/ntaneet/online/candidatelogin.aspx')
    write(app_no, into='पंजीकरण संख्या (Application No.): *')
    write(passwd, into="पासवर्ड (Password) : *")
    click("login")
    return browser

#CBSE result
def get_cbse_data(df, i, *args, **kwargs):
    roll_no = get_prop(df, i, "Roll Number")
    school_no = get_prop(df, i, "School Number")
    centre_no = get_prop(df, i, "Centre Number")
    admit_card_id = get_prop(df, i, "Get Card Id")
    browser = start_chrome('https://cbseresults.nic.in/class12/Class12th20.htm') #class 12
    write(roll_no, into="Enter your Roll Number")
    write(school_no, into="Enter School No.")
    write(centre_no, into="Enter Centre No.")
    write(admit_card_id, into="Enter Admit Card ID.")
    click("Submit")
    wait_until(Text('Check Another Result').exists)
    tables = browser.find_elements_by_tag_name('table')
    candidateInfoTable = tables[4]
    marksTable = tables[5]
    cit_row_elements = candidateInfoTable.find_elements_by_tag_name('tr')
    info = {}
    for element in cit_row_elements:
        tds = element.find_elements_by_tag_name('td')
        [key, value] = [ td.text  for td in tds ]
        info[key] = value
    marks_trs = marksTable.find_elements_by_tag_name('tr')
    marks_lists = [[td.text.strip() for td in tr.find_elements_by_tag_name('td')] for tr in marks_trs]
    marks_df = pd.DataFrame(np.array(marks_lists[1:9]), columns = marks_lists[0])
    result = marks_lists[9][1].strip('Result :').strip()
    #print(info)
    info_1 = {}
    info_1 = marks_df.to_dict()
    for i in range(0, 8):
        info['Sub Code'+str(i)] = info_1['SUB CODE'][i]
        info['Sub Name'+ str(i)] = info_1['SUB NAME'][i]
        info['Theory'+ str(i)] = info_1['THEORY'][i]
        info['Practical' + str(i)] = info_1['PRACTICAL'][i]
        info['Marks' + str(i)] = info_1['MARKS'][i]
        info['Positional Grade' + str(i)] = info_1['POSITIONAL GRADE'][i]
    
    info['Result'] = result
    #print(marks_df)
    #print(info_1)
    #info.update(info_1)
    #print(result)
    #print(info)
    #dfs = pd.read_html(browser.page_source)
    #print(dfs)

    return info, browser
    

#Rsult MAHARASHTRA
def get_maharashtra_data(df, i, *args, **kwargs):
    roll_no = get_prop(df, i, "Roll Number")
    mothers_first_name = get_prop(df, i, "Mother First Name")
    browser = start_chrome('https://mahresult.nic.in/hscresmar20/hscresmar20.htm')
    write(roll_no, into="Roll Number:*")
    write(mothers_first_name, into="Mother's First Name:*")
    click("View Result")
    df = pd.read_html(browser.page_source)
    info_1 = df[0].to_dict()
    stud_info = browser.find_element_by_xpath('/html/body/div/div[3]/div[1]')
    stud_info_1 = stud_info.text
    li = str(stud_info_1).rstrip().split(':')
    li_1 = li[1].rstrip().split('\n')
    li_2 = li[2].rstrip().split('\n')
    info = {}
    info[li[0]] = li_1[0]
    info[li_1[1]] = li_2[0]
    info[li_2[1]] = li[3]

    for i in range(0,7):
        info['Subject Code'+str(i)] = info_1['Subjects Code'][i]
        info['Subject NAme'+str(i)] = info_1['Subject Name'][i]
        info['Marks Obtained'+str(i)] = info_1['Marks Obtained'][i]
    info['Total Marks'] = info_1['Marks Obtained'][7]
    info['Result'] = info_1['Marks Obtained'][8].strip('RESULT : ').strip()

    return info, browser



#KVPY RESULT
def get_kvpy_result(df, i, *args, **kwargs):
    app_no = get_prop(df, i, "Application No.")
    dob = get_prop(df, i, "Date of Birth")
    date = dob.split('-')
    day = date[0]
    month = date[1]
    year = date[2]
    browser = start_chrome('http://kvpy.iisc.ernet.in/kvpy1920/checkMarks.php?msg=Invalid%20Application%20Number%20or%20Date%20Of%20Birth')
    write(app_no, into="Application number:")
    browser.find_element_by_name('dd').send_keys(day)
    browser.find_element_by_name('mm').send_keys(month)
    browser.find_element_by_name('yyyy').send_keys(year)
    click('Continue')
    li = browser.find_element_by_xpath('/html/body/div/div[4]').text
    li_1 = li.split(':')
    li_2 = li_1[1].split('\n')
    li_3 = li_1[2].split('\n')
    li_4 = li_1[3].split('\n')
    li_5 = li_1[4].split('\n')
    info = {}
    info['Application Number'] = li_2[0]
    info[li_2[1]] = li_3[0]
    info[li_3[1]] = li_4[0]
    info[li_4[1]] = li_5[0]
    info[li_5[1]] = li_1[5]

    return info, browser


#mains 2021
def fill_mains_details(df, i, *args, **kwargs):
    candidate_name = get_prop(df, i, "Candidate's Name")
    fathers_name = get_prop(df, i, "Father's Name")
    mothers_name = get_prop(df, i, "Mother's Name")
    date_of_birth = get_prop(df, i, "Date of Birth")
    print(date_of_birth)
    gender = get_prop(df, i, "Gender")
    identity_type = get_prop(df, i, "Identity")
    id_number = get_prop(df, i, "Id Number")
    address = get_prop(df, i, "Address")
    locality = get_prop(df,i, "Locality")
    city_town_village = get_prop(df, i, "City Town Village")
    country = get_prop(df, i, "Country")
    state = get_prop(df, i, "State")
    district = get_prop(df, i, "District")
    pin_code = get_prop(df, i, "Pin Code")
    alternate_contact_number = get_prop(df, i, "Alternate Contact Number")
    email = get_prop(df, i, "Email")
    mobile  = get_prop(df, i, "Mobile Number")
    same_permanent_address = get_prop(df, i, "Permanent Address")
    password = get_prop(df, i, "Password")
    confirm_password = get_prop(df, i, "Password")
    security_Q = get_prop(df, i, "Sequrity Question")
    security_answer = get_prop(df, i, "Sequrity Answer")
    browser = start_chrome('https://jeemain.nta.nic.in/webinfo2021/Page/Page?PageId=1&LangId=P')
    browser.find_element_by_class_name('btn-lg').click()
    browser.switch_to.window(browser.window_handles[1])
    browser.find_element_by_id('ctl00_ContentPlaceHolder1_btnApplyForm').click()
    browser.find_element_by_id('ctl00_ContentPlaceHolder1_chkagreement').click()
    browser.find_element_by_id('ctl00_ContentPlaceHolder1_btnsubmit').click()
    write(candidate_name, into="Candidate's Name")
    write(fathers_name, into="Father's Name")
    write(mothers_name, into="Mother's Name")

    # day, month, year = date_of_birth.split("-")
    # day_menu = browser.find_elements_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_ddlday"]')
    # day_menu.send_keys(day)
    # month_options = {"01": "January (01)", "02": "Februrary (02)", "03": "March (03)", "04": "April (04)", 
    #                  "05": "May (05)", "06": "June (06)", "07": "July (07)", "08": "August (08)", 
    #                  "09": "September (09)", "10": "October (10)", "11": "November (11)", "12": "December (12)"}
    
    # month_menu = browser.find_elements_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_ddlmonth"]')
    # month_menu.send_keys(month_options[month])
    # year_menu = browser.find_elements_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_ddlyear"]')
    # year_menu.send_keys(year)
    
    gender_menu = browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_ddlGender"]')
    gender_menu.send_keys(gender)
    i_type = browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_ddltypeofidentification"]')
    i_type.send_keys(identity_type)
    write(id_number, into="Enter PAN Number")
    address_menu = browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_txtAddress"]')
    address_menu.send_keys(address)
    locality_menu = browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_txtlocality"]')
    locality_menu.send_keys(locality)
    city_menu = browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_txtCity"]')
    city_menu.send_keys(city_town_village)
    browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_ddlCountryAddress_chosen"]/a/span').click()
    country_menu = browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_ddlCountryAddress_chosen"]/div/div/input')
    country_menu.send_keys(country)
    browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_ddlCountryAddress_chosen"]/div/ul/li[1]').click()
    print(state)
    # browser.implicitly_wait(100)
    # browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_ddlStateAddress_chosen"]/a/span').click()
    # state_menu = browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_ddlStateAddress_chosen"]/div/div/input')
    # state_menu.send_keys(state)
    # browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_ddlStateAddress_chosen"]/div/ul/li').click()


    # district_menu = browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_ddlDistrictAddress_chosen"]/div/div/input')
    # district_menu.send_keys(district)

    # pincode_menu = browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_txtPincode"]')

    # pincode_menu.send_keys(pin_code)
    
    # write(pin_code, into="Pin Code")
    print(alternate_contact_number)
    print(email)
    contact_key = browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_txtphoneno"]')
    contact_key.send_keys(alternate_contact_number)
    # write(alternate_contact_number, into="Alternate Contact No. (Optional)")
    email_key = browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_txteid"]')
    email_key.send_keys(email)
    # write(email, into="Email Address")
    mob_menu = browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_txtmobileno"]')
    mob_menu.send_keys(mobile)

    if same_permanent_address == 'TRUE':
        chk_box = browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_chkSameAsCorrespondenceAddress"]')
        click(chk_box)
    else:
        address_P = get_prop(df, i, "Address P")
        locality_P = get_prop(df,i, "Locality P")
        city_town_village_P = get_prop(df, i, "City Town Village P")
        country_p = get_prop(df, i, "Country P")
        state_p = get_prop(df, i, "State P")
        district_p = get_prop(df, i, "District P")
        pin_code_p = get_prop(df, i, "Pin Code P")
        address_menu_1 = browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_txtAddressCoros"]')
        address_menu_1.send_keys(address_P)
        locality_menu_1 = browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_txtlocalityCoros"]')
        locality_menu_1.send_keys(locality_P)
        city_menu_1 = browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_txtCityCoros"]')
        city_menu_1.send_keys(city_town_village_P)
        country_menu_1 = browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_ddlCountryAddressCoros_chosen"]/div/div/input')
        country_menu_1.send_keys(country_p)
        state_menu_1 = browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_ddlStateAddressCoros_chosen"]/div/div/input')
        state_menu_1.send_keys(state_p)
        district_menu_1 = browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_ddlDistrictAddressCoros_chosen"]/div/div/input')
        district_menu_1.send_keys(district_p)
        pincode_menu_1 = browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_txtPincodeCoros"]')
        pincode_menu_1.send_keys(pin_code_p)
    
    write(password, into="Password")
    write(confirm_password, into="Confirm Password")
    seq_q_menu = browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_ddlSecurityQuestion"]')
    seq_q_menu.send_keys(security_Q)
    write(security_answer, into="Security Answer")

    


    return browser




def get_neet_conf_data(df, i, *args, **kwargs):
    browser = login_neet_conf_website(df, i)
    wait_until(Text('Print Confirmation Page').exists)
    # click("Print Confirmation Page")
    click(S('//*[@id="ticker-menu"]/li[2]/ul/li[5]/a'))
    wait_until_img_loads(browser, '//*[@id="Image3"]')
    wait_until_img_loads(browser, '//*[@id="Image2"]')
    wait_until_img_loads(browser, '//*[@id="Image1"]')
    wait_until(Text('Note : - Kindly check following details in Confirmation Page :').exists)

    data = {}
    dfs = pd.read_html(browser.page_source)
    for i, row in dfs[0].loc[2:,:].iterrows():
         data = neet_conf_info(row[0], row[1], i, data)
         data = neet_conf_info(row[2], row[3], i, data)
    return data, browser



def get_neet_omr_data(df, i, *args, **kwargs):
    browser = login_neet_conf_website(df, i)
    wait_until(Text('View OMR').exists)
    go_to('https://ntaneet.nic.in/ntaneet/OMRDisplay/omrdiss.aspx')
    time.sleep(3)
    wait_until_img_loads(browser, '//*[@id="Image1"]')
    zoom_to(browser, 90)
    time.sleep(0.5)
    return {}, browser


def get_wbjee_data(app_no, dob, name, img_path, *args, **kwargs):
    url = 'https://examinationservices.nic.in/examsys/root/AuthForAdmitCardDwd.aspx?enc=WPJ5WSCVWOMNiXoyyomJgNWC6DimOLF6j+KSA+JXppYP01jpy3ljY8ngq1qQmbtA'
    from selenium import webdriver

    chrome_options = webdriver.ChromeOptions()
    prefs = {'download.default_directory' : str(img_path.absolute())}
    chrome_options.add_experimental_option('prefs', prefs)
    # options.add_argument(f"download.default_directory={str(}")

    browser = start_chrome(url, options=chrome_options)
    wait_until_img_loads(browser, '//*[@id="ctl00_ContentPlaceHolder1_captchaimage"]')

    write(app_no, into='Application No :')

    day, month, year = dob.split('-')
    S('//*[@id="ctl00_ContentPlaceHolder1_ddlday"]').web_element.send_keys(day)
    S('//*[@id="ctl00_ContentPlaceHolder1_ddlmonth"]').web_element.send_keys(month)
    S('//*[@id="ctl00_ContentPlaceHolder1_ddlyear"]').web_element.send_keys(year)

    wait_until(Text('Download Rank Card').exists, timeout_secs=50)
    click(S('//*[@id="ctl00_LoginContent_lnkConfirm"]')) #Click download confirmation page
    time.sleep(6)
    conf_page = [i for i in list(img_path.iterdir()) if i.name.startswith(f'ConfirmationPage_{app_no}')][0]
    shutil.move(conf_page, img_path/f'{name}_{app_no}_{dob}_conf_page.pdf')

    click('Download Rank Card')
    wait_until_img_loads(browser, '//*[@id="ctl00_LoginContent_ImgPhoto"]')
    wait_until(Text('(Chairman, WBJEEB)').exists)

    dfs = pd.read_html(browser.page_source)
    data = {}

    df = dfs[2]
    df.columns = [0, 1, 2, 3]
    combine_dicts(data, df[[0, 1]].set_index(0).to_dict()[1])
    combine_dicts(data, df[[2, 3]].set_index(2).to_dict()[3])

    tdata = {}
    df = dfs[3].dropna()[[1, 2, 3]]
    for o in list(df.iloc[0]) + list(df.iloc[1]):
        k, v = o.split(': ')
        tdata[k] = v
    combine_dicts(data, tdata)

    df = dfs[4]
    df.columns = [0, 1, 2, 3]
    combine_dicts(data, df.drop(0)[[0, 1]].set_index(0).to_dict()[1], key_pref='Engg-')
    combine_dicts(data, df.drop(0)[[0, 2]].set_index(0).to_dict()[2], key_pref='Pharm-')

    df = dfs[5]
    df.columns = [0, 1, 2, 3]
    combine_dicts(data, df.iloc[[0, 1], [2, 3]].set_index(2).to_dict()[3])

    
    zoom_to(browser, 70)
    return data, browser


def get_comedk_data(app_no, name, df, i, img_path, *args, **kwargs):
    passwd = get_prop(df, i, 'Password')

    url = 'https://cdn.digialm.com/EForms/configuredHtml/1022/64091/login.html'
    from selenium import webdriver

    chrome_options = webdriver.ChromeOptions()
    prefs = {'download.default_directory' : str(img_path.absolute())}
    chrome_options.add_experimental_option('prefs', prefs)
    browser = start_chrome(url, options=chrome_options)


    wait_until(Text('login').exists)
    time.sleep(2)
    write(app_no, into='ENTER USER ID')
    write(passwd, into='ENTER PASSWORD')
    click('login')
    time.sleep(1)
    browser.set_window_size(1920, 1280)
    if Text('* The login id and password combination does not match.').exists():
        print(f'The Application No. ({app_no}) and password ({passwd}) do not match')
        return {}, browser
    click('COMEDK Rank Card')
    wait_until(Text('here').exists)
    time.sleep(1)
    click('here')
    # click('here')
    time.sleep(6)
    shutil.move(img_path/f'{app_no}.pdf', img_path/f'Rank_card_{name}_{app_no}.pdf')

    for w in browser.window_handles:
        if(w!=browser.current_window_handle):
            browser.switch_to.window(w)
            break

    time.sleep(0.3)
    click('Application View Pdf')
    wait_until(Text('here').exists)
    # click(S('//*[@id="table-32"]/tbody/tr[2]/td/a'))
    time.sleep(2)
    click('here')
    time.sleep(7)
    # print('Application view downloaded')
    shutil.move(img_path/f'{app_no}.pdf', img_path/f'Application_{name}_{app_no}.pdf')

    import tabula
    dfs = tabula.read_pdf(str(img_path/f'Rank_card_{name}_{app_no}.pdf'))
    # print('REad dfs')
    data = {}
    data['TAT No.'] = dfs[0].iloc[0, 1]
    data['STREAM'] = dfs[0].iloc[1, 0].split(' : ')[1]
    # print(data)
    data['Postal Address'] = dfs[1].iloc[0, 1]
    data["Candidate's Name"] = dfs[1].columns[1]
    data['Phone No.'] = dfs[1].iloc[2, 1].split(' ')[0]
    data['Category'] = " ".join(dfs[1].iloc[2, 1].split(' ')[-3:]).lstrip(':')
    data['Date of Birth'] = dfs[1].iloc[3, 1].split(' ')[0]
    data['HK Region'] = dfs[1].iloc[3, 1].split(' ')[-1]
    data['Linguistic Minority'] = dfs[1].iloc[4, 1].split(' ')[0]
    data['Religious Minority'] = " ".join(dfs[1].iloc[4, 1].split(' ')[-2:]).lstrip(':')
    # print(data)
    data['Physics'] = dfs[2].iloc[2, 2]
    data['Chemistry'] = dfs[2].iloc[3, 2]
    data['Mathematics'] = dfs[2].iloc[4, 2]
    data['SCORE'] = dfs[2].iloc[5, 2]
    data['Rank'] = dfs[2].iloc[8, 1]
    # print(data)
    data['Physics-in words'] = dfs[2].iloc[2, 3]
    data['Chemistry-in words'] = dfs[2].iloc[3, 3]
    data['Mathematics-in words'] = dfs[2].iloc[4, 3]
    data['SCORE-in words'] = dfs[2].iloc[5, 3]
    data['Rank-in words'] = dfs[2].iloc[8, 2]
    print("Done")
    browser.switch_to.window(browser.window_handles[0])
    return data, browser


