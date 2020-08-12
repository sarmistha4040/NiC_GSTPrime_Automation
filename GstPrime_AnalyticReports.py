from selenium import webdriver
from selenium.common.exceptions import NoAlertPresentException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.wait import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import datetime
import openpyxl
import time
import os


Curr_Return_Month = "June"
Curr_Coll_Month = "July"
Curr_Year = "2020"
Grp_range = "10000"
#Mid_range = "75K - 2.5L"
Mid_range = "30K - 75K"
#Other_range = "2.5L - 8L "
Other_range = "2.5L - 6L "
curr_month_start = '01/07/2020'
curr_month_end = '31/07/2020'
curr_year = datetime.datetime.now()
year = curr_year.year




#******CBIC Credentials ;
# uid = "zonecbic"
# pwd = "Abc123@@"
# url = 'http://gst.kar.nic.in/gstprime_cbic'

#****** ..... JH Credentials ;
# url = 'http://10.92.240.137/gstprimenew/'
# uid = 'gstnic'
# pwd = 'nic@6239'

#*** ***...... ASSAM
#uid = "hqtest"
#uid = "zonetest"
#uid = "divtest"
#uid = "fotest"
#pwd = "Nic@1234"
#url = 'http://103.8.248.152/gstprimenew'
#state='ASM'


#*** URL and credentials for Maharastra
# uid = 'gstnic'
#uid = 'hqtest'
# uid = 'divtest'
#uid = 'fotest'
# pwd = 'nic@6239'
#  pwd = 'Nic@1234'
#  url = 'http://172.30.248.23/gstprime_test/'
#  state='MH'


# *** URL and credentials for West Bengal  GST Prime_test
# uid = 'hqtest'
# pwd = 'Nic@1234'
# url = 'http://10.173.57.166/gstprime_test'
# state = 'WB'





# *** URL and credentials for BIHAR  GST Prime
uid = 'hqtest'
#uid = 'divtest'
# uid = 'fotest'
pwd = 'Nic@1234'
url = 'http://164.100.130.227/gstprime/'
state = 'Bihar'



try:
    # .....Media setup for storing images/screenshots
    cwd = os.getcwd()  # get current directory path
    parent_dir = str(cwd)
    directory = "Media" # create new directory
    path = os.path.join(parent_dir, directory)
    isdir = os.path.isdir(path)
    if not isdir:
        os.mkdir(path)
        print("Media Directory created")
    else:
        print("Media Directory found, processing to save files!")

    # .....Fetching Data into Excel sheet
    driver=webdriver.Chrome()

    ###Open the Template Excel file
    srcfile = openpyxl.load_workbook('GstPrime_AnalyticReports_inputvalues_Sheet.xlsx', read_only=False,
                                 keep_vba=True)
    sheetname = srcfile['input values']
    error_sheet = srcfile['Errors']
    sheetname['c2'] = url + " " + str(datetime.datetime.now())
    print("----------------\nURL: " + url + "\n------------------")
    # .....
    driver.get(url)
    username = driver.find_element(By.ID, 'txt_username').send_keys(uid)
    password = driver.find_element(By.ID, 'txt_password').send_keys(pwd)
    submit = driver.find_element(By.XPATH, "//input[@type='submit']").click()
    time.sleep(3)
    sheetname['c3'] = "username::" + uid
    sheetname['c4'] = "password::" + pwd

    # .....
    # Home Page

    Page_title = driver.title
    if Page_title == 'GSTPrime Main Menu page':
        print(Page_title)
    else:
         destination = "Media//homepage.png"
         driver.save_screenshot(destination)
         print("Error Screenshot saved as ::", destination)
         print("Done: Homepage !")
         time.sleep(3)

    # 1.......... Analytic-Filings-R3B comp with Avg.
    # .....Top...
    # def r3bfilers():
    driver.get(url + "/Reports/rptr3btoptaxpayers.aspx")
    time.sleep(5)
    month = driver.find_element(By.ID, 'ctl00_ddl_Month').send_keys(Curr_Return_Month)

    range = driver.find_element(By.ID, 'ctl00_ddl_tax_payer_range').send_keys(Grp_range)
    year = driver.find_element(By.ID, "ctl00_ddl_year").send_keys(Curr_Year)
    time.sleep(3)
    driver.find_element(By.XPATH, "//input[@id='ctl00_btn_go']").click()
    time.sleep(5)
    page_title = driver.title
    sheetname['a4'] = "Analytic_Filings_R3B Filers_R3B comp with avg -Top"
    sheetname['b5'] = "Month::" + "" + Curr_Return_Month
    sheetname['b6'] = "Grp range::" + "" + Grp_range
    sheetname['b7'] = "Year::" + "" + Curr_Year
    try:
        element = driver.find_element(By.XPATH, '//table[@id="grd_main"]')
        WebDriverWait(driver, 20).until(EC.visibility_of((element)))
        boolean = element.is_displayed()
        print(boolean)
        grd_main = driver.find_element(By.XPATH, '//table[@id="grd_main"]')
        grid_rows = grd_main.find_elements_by_tag_name("tr")
        row_count = len(grid_rows)
        if (page_title == 'Analysis report on GSTR3B for Top Filers' or row_count>2):

            Total_Count = driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_lblcount').text
            print("::" + Total_Count )

            sheetname['b4'] = "R3B comp with average Top :: Loading fine !"


        elif(row_count == 1 ):
            #print("R3B comp with average Top-:: No data found")
            error_sheet['b1']= "R3B comp with average Top"
            error_sheet['c1'] = " No data found"
            error_sheet['b2'] = "Month::" + "" + Curr_Return_Month
            error_sheet['b3'] = "Grp range::" + "" + Grp_range
            error_sheet['b4'] = "Year::" + "" + Curr_Year
    except Exception as e:
        print("The Exception found:", str(e))
        destination = "Media//R3BFilingsComp_WithAvg_top.png"
        driver.save_screenshot(destination)
        print("Error Screenshot saved as ::", destination)
        print("The Exception found:", str(e))
        error_sheet['b1'] = "R3B comp with average Top"
        error_sheet['c1'] = "Error page occurs"
        error_sheet['b2'] = "Month::" + "" + Curr_Return_Month
        error_sheet['b3'] = "Grp range::" + "" + Grp_range
        error_sheet['b4'] = "Year::" + "" + Curr_Year

    # middle...
    driver.get(url + "/Reports/rptr3btoptaxpayers.aspx")
    Middle = '//input[@id="ctl00_rdb_taxpayer_type_1"]'
    driver.find_element(By.XPATH, Middle).click()
    time.sleep(3)
    month = driver.find_element(By.ID, 'ctl00_ddl_Month').send_keys(Curr_Return_Month)
    range = driver.find_element(By.ID, 'ctl00_ddl_tax_payer_range').send_keys(Mid_range)
    year = driver.find_element(By.ID, 'ctl00_ddl_year').send_keys(Curr_Year)
    element = driver.find_element(By.ID, 'ctl00_btn_go')
    element.send_keys(Keys.ENTER)
    time.sleep(5)
    sheetname['b9'] = "Month::" + "" + Curr_Return_Month
    sheetname['b10'] = "Grp range::" + "" + Mid_range
    sheetname['b11'] = " Year::" + "" + Curr_Year
    try:
        element = driver.find_element(By.XPATH, '//*[@id="grd_main"]')
        WebDriverWait(driver, 20).until(EC.visibility_of((element)))
        if page_title == 'Analysis report on GSTR3B for Top Filers':
            sheetname['b8'] = "R3B comp with average middle :: Loading fine !"
            sheetname['a8'] = "R3B comp with average middle "
            grd_main = driver.find_element(By.XPATH, '//*[@id="grd_main"]')
            grid_rows = grd_main.find_elements_by_tag_name("tr");
            row_count = len(grid_rows)
            Total_Count = driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_lblcount').text
            print("::" +Total_Count)
        elif (row_count == 1):
            error_sheet['b5'] = "R3B comp with average Middle"
            error_sheet['c5'] = " No data found"
            error_sheet['b6'] = "Month::" + "" + Curr_Return_Month
            error_sheet['b7'] = "Grp range::" + "" + Mid_range
            error_sheet['b8'] = " Year::" + "" + Curr_Year

    except Exception as e:
        print("The Exception found:", str(e))
        destination = "Media//R3BFilingsCompWithAvg_Middle.png"
        driver.save_screenshot(destination)
        print("Error Screenshot saved as ::", destination)
        error_sheet['b5'] = "R3B comp with average Middle"
        error_sheet['c5'] = " No data found"
        error_sheet['b6'] = "Month::" + "" + Curr_Return_Month
        error_sheet['b7'] = "Grp range::" + "" + Mid_range
        error_sheet['b8'] = " Year::" + "" + Curr_Year

        # Others.....
        Others = '//input[@id="ctl00_rdb_taxpayer_type_2"]'
        driver.find_element(By.XPATH, Others).click()
        Month = 'ctl00_ddl_Month'
        month = driver.find_element(By.ID, Month).send_keys(Curr_Return_Month)
        Range = 'ctl00_ddl_tax_payer_range'
        range = driver.find_element(By.ID, Range).send_keys(Other_range)
        Year = 'ctl00_ddl_year'
        year = driver.find_element(By.ID, Year).send_keys(Curr_Year)
        Go = 'ctl00_btn_go'
        element = driver.find_element(By.ID, Go)
        element.send_keys(Keys.ENTER)
        time.sleep(5)
        sheetname['b13'] = "Month::" + "" + Curr_Return_Month
        sheetname['b14'] = "Grp range::" + "" + Other_range
        sheetname['b15'] = " Year::" + "" + Curr_Year
    try:
        element = driver.find_element(By.XPATH, '//*[@id="grd_main"]')
        WebDriverWait(driver, 20).until(EC.visibility_of((element)))
        if  page_title == 'Analysis report on GSTR3B for Top Filers':
            grd_main = driver.find_element(By.XPATH, '//*[@id="grd_main"]')
            grid_rows = grd_main.find_elements_by_tag_name("tr");
            row_count = len(grid_rows)
            Total_Count = driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_lblcount').text
            print("::" +Total_Count )
            sheetname['b12'] = "R3B comp with average other :: Loading fine !"
            sheetname['a12'] = "R3B comp with average other "
            sheetname['b13'] = "Month::" + "" + Curr_Return_Month
            sheetname['b14'] = "Grp range::" + "" + Other_range
            sheetname['b15'] = " Year::" + "" + Curr_Year
        elif (row_count == 1 ):
            error_sheet['b9'] = "R3B comp with average Other"
            error_sheet['c9'] = " No data found"
            error_sheet['b10'] = "Month::" + "" + Curr_Return_Month
            error_sheet['b11'] = "Grp range::" + "" + Other_range
            error_sheet['b12'] = " Year::" + "" + Curr_Year

    except Exception as e:
            print("The Exception found:", str(e))
            destination = "Media//R3BCompWithAvg_other.png"
            driver.save_screenshot(destination)
            print("Error Screenshot saved as ::", destination)
            print("The Exception found:", str(e))
            error_sheet['b9'] = "R3B comp with average Other"
            error_sheet['c9'] = " No data found"
            error_sheet['b10'] = "Month::" + "" + Curr_Return_Month
            error_sheet['b11'] = "Grp range::" + "" + Other_range
            error_sheet['b12'] = " Year::" + "" + Curr_Year

    finally:
            print("Done : rptr3btopTaxpayers_R3bCompWithAvg.....")


    # 2 #.......... Analytic-Filings-R3B Comp Month On Month
    # Analytic--> Filings--> R3B_Comp_Month On Month--> r3btaxpayercomp
    driver.get(url + "/Reports/r3btaxpayercomp.aspx")
    time.sleep(5)
    Month = 'ctl00_MainContent_frm_ddl_Month'
    driver.find_element(By.ID,Month).send_keys(Curr_Return_Month)
    Year = 'ctl00_MainContent_frm_ddl_year'
    driver.find_element(By.ID,Year).send_keys(Curr_Year)
    Range= "ctl00_MainContent_ddl_tomg"
    driver.find_element(By.ID,Range).send_keys(Grp_range)
    time.sleep(5)
    Go="ctl00_MainContent_btnshow"
    element = driver.find_element(By.ID,Go)
    element.send_keys(Keys.ENTER)
    page_title=driver.title
    sheetname['a16'] = "R3BTaxpayerComp_Filings_MonthNMonth_Top"
    sheetname['b17'] = "Month::"+ "" + Curr_Return_Month
    sheetname['b18'] = " Year::" + " "+ Curr_Year
    sheetname['b19'] = "Grp_range" + " " +Grp_range
    try:
        #grid_main="//table[@id='ctl00_MainContent_grd_prd_ws']"
        grid_main="//table[@id='grd_prd_ws']"
        Comp_Month_Month = driver.find_element(By.XPATH, grid_main)
        WebDriverWait(driver, 20).until(EC.visibility_of((Comp_Month_Month)))
        #page_title == 'Analysis report on comparative tax payment by top tax payers':
        grd_main = driver.find_element(By.XPATH, grid_main)
        grid_rows = grd_main.find_elements_by_tag_name("tr");
        row_count = len(grid_rows)
        if (row_count>= 2 ):
            sheetname['b16'] = "R3BComp-monthNmonthTop :: Loading fine !"

        elif (row_count ==1):
            error_sheet['b13'] = "R3B Comp MonthNMonthTop"
            error_sheet['c13'] = "No data found"
            error_sheet['b14'] = "Month::" + "" + Curr_Return_Month
            error_sheet['b15'] = "Grp range::" + "" + Grp_range
            error_sheet['b16'] = " Year::" + "" + Curr_Year

    except Exception as e:
        print("The Exception found:", str(e))
        destination = "Media//R3BCompMonthNMonth.png"
        driver.save_screenshot(destination)
        print("Error Screenshot saved as ::", destination)
        error_sheet['b13'] = "R3B Comp MonthNMonthTop"
        error_sheet['c13'] = "No data found"
        error_sheet['b14'] = "Month::" + "" + Curr_Return_Month
        error_sheet['b15'] = "Grp range::" + "" + Grp_range
        error_sheet['b16'] = " Year::" + "" + Curr_Year
    # middle...
    driver.get(url + "/Reports/r3btaxpayercomp.aspx")
    time.sleep(7)
    element=driver.find_element(By.XPATH,"//input[@id='ctl00_MainContent_rdb_taxpayer_type_1']")
    element.send_keys(Keys.ENTER)
    driver.find_element(By.ID,'ctl00_MainContent_frm_ddl_Month').send_keys(Curr_Return_Month)
    time.sleep(2)
    driver.find_element(By.ID,'ctl00_MainContent_frm_ddl_year').send_keys(Curr_Year)
    time.sleep(2)
    Range= "ctl00_MainContent_ddl_tomg"
    driver.find_element(By.ID,Range).send_keys(Mid_range)
    Go="ctl00_MainContent_btnshow"
    element = driver.find_element(By.ID,Go)
    element.send_keys(Keys.ENTER)
    sheetname['a21'] = "R3BComp-monthNmonthMiddle "
    sheetname['b22'] = "Month::"+ "" + Curr_Return_Month
    sheetname['b23'] = " Year::" + " "+ Curr_Year
    sheetname['b24'] = "Grp_range" + " "+ Mid_range
    try:
        grid_main="//table[@id='grd_prd_ws']"
        Comp_Month_Month = driver.find_element(By.XPATH, grid_main)
        WebDriverWait(driver, 20).until(EC.visibility_of((Comp_Month_Month)))
        # page_title = driver.title
        # if  page_title == 'Analysis report on comparative tax payment by top tax payers':
        grd_main = driver.find_element(By.XPATH, grid_main)
        grid_rows = grd_main.find_elements_by_tag_name("tr");
        row_count = len(grid_rows)
        if (row_count >=2):
            sheetname['b21'] = "R3BComp-monthNmonthMiddle :: Loading fine !"
            sheetname['a21'] = "R3BComp-monthNmonthMiddle "
        elif (row_count == 1 ):
            error_sheet['b17'] = "R3B Comp MonthNMonth Middle"
            error_sheet['c17'] = "No data found"
            error_sheet['b18'] = "Month::" + "" + Curr_Return_Month
            error_sheet['b19'] = "Grp range::" + "" + Mid_range
            error_sheet['b20'] = " Year::" + "" + Curr_Year
    except Exception as e:
            print("The Exception found:", str(e))
            destination = "Media//R3BTaxpayerCompMid.png"
            driver.save_screenshot(destination)
            print("Error Screenshot saved as ::", destination)
            error_sheet['b17'] = "R3B Comp MonthNMonth Middle"
            error_sheet['c17'] = "Error page occurs"
            error_sheet['b18'] = "Month::" + "" + Curr_Return_Month
            error_sheet['b19'] = "Grp range::" + "" + Mid_range
            error_sheet['b20'] = " Year::" + "" + Curr_Year
    # Others...
    time.sleep(7)
    driver.get(url + "/Reports/r3btaxpayercomp.aspx")
    time.sleep(4)
    element=driver.find_element(By.XPATH,"//input[@id='ctl00_MainContent_rdb_taxpayer_type_2']")
    element.send_keys(Keys.ENTER)
    Month = 'ctl00_MainContent_frm_ddl_Month'
    driver.find_element(By.ID,Month).send_keys(Curr_Return_Month)
    Year = 'ctl00_MainContent_frm_ddl_year'
    driver.find_element(By.ID,Year).send_keys(Curr_Year)
    time.sleep(2)
    Range = "ctl00_MainContent_ddl_tomg"
    driver.find_element(By.ID,Range).send_keys(Other_range)
    time.sleep(5)
    Go="ctl00_MainContent_btnshow"
    element = driver.find_element(By.ID,Go)
    element.send_keys(Keys.ENTER)
    time.sleep(5)

    sheetname['a26'] = "R3BComp-monthNmonthOthers "
    sheetname['b27'] = "Month::"+ "" + Curr_Return_Month
    sheetname['b28'] = " Year::" + " "+ Curr_Year
    sheetname['b29'] = "Grp_range" + " " +Other_range
    try:
        grid_main="//table[@id='grd_prd_ws']"
        Comp_Month_Month = driver.find_element(By.XPATH, grid_main)
        WebDriverWait(driver, 20).until(EC.visibility_of((Comp_Month_Month)))
        if  page_title == 'Analysis report on comparative tax payment by top tax payers':
            grd_main = driver.find_element(By.XPATH, grid_main)
            grid_rows = grd_main.find_elements_by_tag_name("tr");
            row_count = len(grid_rows)
            sheetname['b26'] = "R3BComp-monthNmonthOthers :: Loading fine !"

        elif (row_count == 1 ):
            error_sheet['b22'] = "R3BComp-monthNmonthOthers"
            error_sheet['c22'] = "No data found"
            error_sheet['b23'] = "Month::" + "" + Curr_Return_Month
            error_sheet['b24'] = "Grp range::" + "" + Other_range
            error_sheet['b25'] = " Year::" + "" + Curr_Year

    except Exception as e:
        print("The Exception found:", str(e))
        destination = "Media//R3BTaxpayerCompOther.png"
        driver.save_screenshot(destination)
        print("Error Screenshot saved as ::", destination)

        error_sheet['b22'] = "R3BComp-monthNmonthOthers"
        error_sheet['c22'] = "Error page occurs"
        error_sheet['b23'] = "Month::" + "" + Curr_Return_Month
        error_sheet['b24'] = "Grp range::" + "" + Other_range
        error_sheet['b25'] = " Year::" + "" + Curr_Year


    finally:
        #print("Done : r3bTaxpayerComp")
        print("Done :r3bTaxpayerComp_R3BTaxPayer Comp MonthNMonth......")

    # 3 #.......... Analytic-Filings-R3B New taxpayer.

    time.sleep(5)
    driver.get(url + "/Reports/Rpt_AllTaxReport.aspx?txt_rpttype=5")
    try:
        time.sleep(3)
        Month = "ctl00_ContentPlaceHolder1_ddl_month1"
        driver.find_element(By.ID, Month).send_keys(Curr_Return_Month)
        Year = "ctl00_ContentPlaceHolder1_ddl_period"
        driver.find_element(By.ID, Year).send_keys(Curr_Year)
        time.sleep(2)
        Go = "ctl00_ContentPlaceHolder1_Button1"
        element = driver.find_element(By.ID, Go)
        element.send_keys(Keys.ENTER)
        time.sleep(7)
        sheetname['a33'] = "R3B New taxpayer"
        sheetname['b34'] = "Month::" + "" + Curr_Return_Month
        sheetname['b35'] = "Year::" + " " + Curr_Year
        page_title = driver.title
        # print(page_title)
        time.sleep(7)
        grid_main = "//table[@id='grd_main']"
        Newtaxpayer = driver.find_element(By.XPATH, grid_main)
        WebDriverWait(driver, 20).until(EC.visibility_of((Newtaxpayer)))
        page_title = 'Report on New Registration'
        grd_main = driver.find_element(By.XPATH, grid_main)
        grid_rows = grd_main.find_elements_by_tag_name("tr");
        row_count = len(grid_rows)
        if (row_count >= 2):
            #Total_Count = driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_lblcount').text
            #print("::" +Total_Count)
            sheetname['b33'] = "R3BNewTaxpayers :: Loading fine !"
            sheetname['a33'] = "R3BNewTaxpayers "
        elif (row_count == 1 ):

            error_sheet['b26'] = "R3BNewTaxpayers"
            error_sheet['c26'] = "No data found "
            error_sheet['b27'] = "Month::" + "" + Curr_Return_Month
            error_sheet['b28'] = " Year::" + "" + Curr_Year

    except Exception as e:
        print("The error occured as:"+ str(e))
        destination = "Media//R3BNewTaxpayers.png"
        driver.save_screenshot(destination)
        print("Error Screenshot saved as ::", destination)
        error_sheet['b26'] = "R3BNewTaxpayers"
        error_sheet['c26'] = "Error page occurs "
        error_sheet['b27'] = "Month::" + "" + Curr_Return_Month
        error_sheet['b28'] = " Year::" + "" + Curr_Year


    finally:
        print("Done : Rpt_AllTaxReport_NewTaxPayer......")

    # 4 #.......... Analytic-Filings-R3B Year on Year.....................

    driver.get(url + "/Reports/rptR3BTopTaxPayrs_PrdComp_byfin.aspx")
    time.sleep(5)
    driver.find_element(By.ID, 'ctl00_ddl_Month').send_keys(Curr_Return_Month)
    driver.find_element(By.ID, 'ctl00_ddl_tax_payer_range').send_keys(Grp_range)
    driver.find_element(By.ID, 'ctl00_ddl_year').send_keys(Curr_Year)
    time.sleep(5)
    sheetname['a36'] = "R3B_Comp_Year_On_Year_Top"
    sheetname['b37'] = "Month::" + "" + Curr_Return_Month
    sheetname['b38'] = "Grp range::" + "" + Grp_range
    sheetname['b39'] = " Year::" + "" + Curr_Year
    try:
        time.sleep(3)
        driver.find_element(By.XPATH, "//input[@id='ctl00_btn_go']").send_keys(Keys.ENTER)
        time.sleep(5)
        grid_main = "//table[@id='grid_main']"
        element = driver.find_element(By.XPATH, "//table[@id='grd_main']")
        WebDriverWait(driver, 20).until(EC.visibility_of((element)))
        grid = driver.find_element(By.XPATH, "//table[@id='grd_main']")
        grid_rows = grid.find_elements_by_tag_name("tr");
        row_count = len(grid_rows)
        if (row_count > 2):
            Total_Count = driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_lblcount').text
            print("::" +Total_Count)
            # print("R3B Comp Year on Year page is loading fine")
            sheetname['b36'] = "R3B Comp Year on Year :: Loading fine !"
            #sheetname['a37'] = "R3B Comp Year on Year "
        elif(row_count == 1):
            error_sheet['b29'] = "R3B Comp Year on Year Top"
            error_sheet['c29'] = "No data found "
            error_sheet['b30'] = "Month::" + "" + Curr_Return_Month
            error_sheet['b31'] = "Grp range::" + "" + Grp_range
            error_sheet['b32'] = " Year::" + "" + Curr_Year
    except Exception as e:
        print("The Exception found:", str(e))
        destination = "Media//R3BYearOnYearTop.png"
        driver.save_screenshot(destination)
        print("Error Screenshot saved as ::", destination)
        error_sheet['b29'] = "R3B Comp Year on Year Top"
        error_sheet['c29'] = "Error page occurs "
        error_sheet['b30'] = "Month::" + "" + Curr_Return_Month
        error_sheet['b31'] = "Grp range::" + "" + Grp_range
        error_sheet['b32'] = " Year::" + "" + Curr_Year

    # middle ....
    driver.get(url + "/Reports/rptR3BTopTaxPayrs_PrdComp_byfin.aspx")
    time.sleep(5)
    driver.find_element(By.XPATH, "//input[@id='ctl00_rdb_taxpayer_type_1']").click()
    driver.find_element(By.ID, 'ctl00_ddl_Month').send_keys(Curr_Return_Month)
    driver.find_element(By.ID, 'ctl00_ddl_tax_payer_range').send_keys(Mid_range)
    driver.find_element(By.ID, 'ctl00_ddl_year').send_keys(Curr_Year)
    driver.find_element(By.ID, 'ctl00_btn_go').send_keys(Keys.ENTER)
    time.sleep(5)
    sheetname['b43'] = "Month::" + "" + Curr_Return_Month
    sheetname['a42'] = "R3B Comp Year on Year middle "
    sheetname['b44'] = "Grp range::" + "" + Mid_range
    sheetname['b45'] = " Year::" + "" + Curr_Year
    try:
        element = driver.find_element(By.XPATH, "//table[@id='grd_main']")
        WebDriverWait(driver, 20).until(EC.visibility_of((element)))
        grid = driver.find_element(By.XPATH, "//table[@id='grd_main']")
        grid_rows = grid.find_elements_by_tag_name("tr");
        row_count = len(grid_rows)
        if (row_count >= 2):
            Total_Count = driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_lblcount').text
            print("::" +Total_Count)
            sheetname['b42'] = "R3B Comp Year on Year middle :: Loading fine !"

        elif (row_count == 1 ):
            error_sheet['b33'] = "R3B Comp Year on Year Middle"
            error_sheet['c33'] = "No data found"
            error_sheet['b34'] = "Month::" + "" + Curr_Return_Month
            error_sheet['b35'] = "Grp range::" + "" + Mid_range
            error_sheet['b36'] = " Year::" + "" + Curr_Year

    except Exception as e:
        print("The Exception found:", str(e))
        destination = "Media//R3BYearOnYearMiddle.png"
        driver.save_screenshot(destination)
        print("Error Screenshot saved as ::", destination)
        error_sheet['b33'] = "R3B Comp Year on Year Middle"
        error_sheet['c33'] = "Error page occurs"
        error_sheet['b34'] = "Month::" + "" + Curr_Return_Month
        error_sheet['b35'] = "Grp range::" + "" + Mid_range
        error_sheet['b36'] = " Year::" + "" + Curr_Year

    # others...
    driver.get(url + "/Reports/rptR3BTopTaxPayrs_PrdComp_byfin.aspx")
    time.sleep(5)
    driver.find_element(By.XPATH, "//input[@id='ctl00_rdb_taxpayer_type_2']").click()
    driver.find_element(By.ID, 'ctl00_ddl_Month').send_keys(Curr_Return_Month)
    driver.find_element(By.ID, 'ctl00_ddl_tax_payer_range').send_keys(Other_range)
    driver.find_element(By.ID, 'ctl00_ddl_year').send_keys(Curr_Year)
    driver.find_element(By.ID, 'ctl00_btn_go').send_keys(Keys.ENTER)
    time.sleep(5)
    sheetname['a46'] = "R3B comp Year on Year other "
    sheetname['b47'] = "Month::" + "" + Curr_Return_Month
    sheetname['b48'] = "Grp range::" + "" + Other_range
    sheetname['b49'] = " Year::" + "" + Curr_Year
    time.sleep(5)
    try:
            grid_main = "//table[@id='grid_main']"
            element = driver.find_element(By.XPATH, "//table[@id='grd_main']")
            WebDriverWait(driver, 20).until(EC.visibility_of((element)))
            grid = driver.find_element(By.XPATH, "//table[@id='grd_main']")
            grid_rows = grid.find_elements_by_tag_name("tr");
            row_count = len(grid_rows)
            if (row_count >= 2):
                Total_Count = driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_lblcount').text
                print("::" +Total_Count)
                sheetname['b46'] = "R3B Year on Year other :: Loading fine !"

            elif (row_count == 1):
                error_sheet['b37'] = "R3B Comp Year on Year Others"
                error_sheet['c37'] = "No data found"
                error_sheet['b38'] = "Month::" + "" + Curr_Return_Month
                error_sheet['b39'] = "Grp range::" + "" + Other_range
                error_sheet['b40'] = " Year::" + "" + Curr_Year
    except Exception as e:
            print("The Exception found:", str(e))
            destination = "Media//R3BYearOnYearOther.png"
            driver.save_screenshot(destination)
            print("Error Screenshot saved as ::", destination)
            error_sheet['b37'] = "R3B Comp Year on Year Others"
            error_sheet['c37'] = "Error page occurs"
            error_sheet['b38'] = "Month::" + "" + Curr_Return_Month
            error_sheet['b39'] = "Grp range::" + "" + Other_range
            error_sheet['b40'] = " Year::" + "" + Curr_Year
    finally:
        print("Done :R3B Year On Year_rptR3BTopTaxPayrs_YearOnYear......")

    # B,1 #.......... Analytics...........Matching............

    driver.get(url + "/Reports/rpt_Mismatch_Riskbased.aspx")
    time.sleep(5)
    Month = 'ctl00_ddl_Month'
    Go = 'ctl00_btn_go'
    Range = 'ctl00_ddl_tax_payer_range'
    Year = 'ctl00_ddl_year'
    driver.find_element(By.ID, Month).send_keys(Curr_Return_Month)
    driver.find_element(By.ID, Range).send_keys(Grp_range)
    driver.find_element(By.ID, Year).send_keys(Curr_Year)
    driver.find_element(By.ID, Go).send_keys(Keys.ENTER)
    time.sleep(5)
    sheetname['a50'] = "Matching - Mismatch Risk based Top"
    sheetname['b51'] = "Month::" + "" + Curr_Return_Month
    sheetname['b52'] = "Grp range::" + "" + Grp_range
    sheetname['b53'] = " Year::" + "" + Curr_Year
    time.sleep(5)
    try:
        grid_main = "//table[@id='ctl00_ContentPlaceHolder1_grd_main']"
        element = driver.find_element(By.XPATH, grid_main)
        WebDriverWait(driver, 20).until(EC.visibility_of((element)))
        grid = driver.find_element(By.XPATH, grid_main)
        grid_rows = grid.find_elements_by_tag_name("tr");
        row_count = len(grid_rows)
        if (row_count >= 2):
            Total_Count = driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_lblcount').text
            print("::" +Total_Count)
            sheetname['b50'] = "Mismatch Risk based Top :: Loading fine !"
            #sheetname['a47'] = "Mismatch Risk based "
        elif (row_count == 1 ):
            error_sheet['b37'] = "Mismatch Risk based Top "
            error_sheet['c37'] = "No data found"
            error_sheet['b38'] = "Month::" + "" + Curr_Return_Month
            error_sheet['b39'] = "Grp range::" + "" + Grp_range
            error_sheet['b40'] = " Year::" + "" + Curr_Year
    except Exception as e:
        print("The Exception found:", str(e))
        destination = "Media//MismatchRiskbasedTop .png"
        driver.save_screenshot(destination)
        print("Error Screenshot saved as ::", destination)
        error_sheet['b37'] = "Mismatch Risk based Top"
        error_sheet['c37'] = "Error page occurs"
        error_sheet['b38'] = "Month::" + "" + Curr_Return_Month
        error_sheet['b39'] = "Grp range::" + "" + Grp_range
        error_sheet['b40'] = " Year::" + "" + Curr_Year
    # Middle ...
    driver.get(url + "/Reports/rpt_Mismatch_Riskbased.aspx")
    time.sleep(5)
    Month = 'ctl00_ddl_Month'
    Go = 'ctl00_btn_go'
    Range = 'ctl00_ddl_tax_payer_range'
    Year = 'ctl00_ddl_year'
    driver.find_element(By.XPATH, "//input[@id='ctl00_rdb_taxpayer_type_1']").click()
    driver.find_element(By.ID, Month).send_keys(Curr_Return_Month)
    driver.find_element(By.ID, Range).send_keys(Mid_range)
    driver.find_element(By.ID, Year).send_keys(Curr_Year)
    driver.find_element(By.ID, Go).send_keys(Keys.ENTER)
    time.sleep(5)
    sheetname['a54'] = "Mismatch Risk based middle"
    sheetname['b55'] = "Month::" + "" + Curr_Return_Month
    sheetname['b56'] = "Grp range::" + "" + Mid_range
    sheetname['b57'] = " Year::" + "" + Curr_Year
    time.sleep(5)
    try:
        grid_main = "//table[@id='ctl00_ContentPlaceHolder1_grd_main']"
        element = driver.find_element(By.XPATH, grid_main)
        WebDriverWait(driver, 20).until(EC.visibility_of((element)))
        grid = driver.find_element(By.XPATH, grid_main)
        grid_rows = grid.find_elements_by_tag_name("tr");
        row_count = len(grid_rows)
        if (row_count >= 2):
            Total_Count = driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_lblcount').text
            print("::" +Total_Count)
            sheetname['b54'] = "Mismatch Risk based middle :: Loading fine !"
            sheetname['a54'] = "Mismatch Risk based middle"
        elif (row_count == 1 ):
            error_sheet['b41'] = "Mismatch Risk based middle "
            error_sheet['c41'] = "No data found "
            error_sheet['b42'] = "Month::" + "" + Curr_Return_Month
            error_sheet['b43'] = "Grp range::" + "" + Mid_range
            error_sheet['b44'] = " Year::" + "" + Curr_Year
    except Exception as e:
        print("The Exception found:", str(e))
        destination = "Media//MismatchRiskbased_middle.png"
        driver.save_screenshot(destination)
        print("Error Screenshot saved as ::", destination)

        error_sheet['b41'] = "Mismatch Risk based middle "
        error_sheet['c41'] = "Error Page Occurs"
        error_sheet['b42'] = "Month::" + "" + Curr_Return_Month
        error_sheet['b43'] = "Grp range::" + "" + Mid_range
        error_sheet['b44'] = " Year::" + "" + Curr_Year

    # Others...
    driver.get(url + "/Reports/rpt_Mismatch_Riskbased.aspx")
    time.sleep(5)
    Month = 'ctl00_ddl_Month'
    Go = 'ctl00_btn_go'
    Range = 'ctl00_ddl_tax_payer_range'
    Year = 'ctl00_ddl_year'
    driver.find_element(By.XPATH, "//input[@id='ctl00_rdb_taxpayer_type_2']").click()
    driver.find_element(By.ID, Month).send_keys(Curr_Return_Month)
    driver.find_element(By.ID, Range).send_keys(Other_range)
    driver.find_element(By.ID, Year).send_keys(Curr_Year)
    driver.find_element(By.ID, Go).send_keys(Keys.ENTER)
    time.sleep(5)
    sheetname['b59'] = "Month::" + "" + Curr_Return_Month
    sheetname['b60'] = "Grp range::" + "" + Other_range
    sheetname['b61'] = " Year::" + "" + Curr_Year
    sheetname['a58'] = "Matching - Mismatch Risk based Other "
    try:
        grid_main = "//table[@id='ctl00_ContentPlaceHolder1_grd_main']"
        element = driver.find_element(By.XPATH, grid_main)
        WebDriverWait(driver, 20).until(EC.visibility_of((element)))
        grid = driver.find_element(By.XPATH, grid_main)
        grid_rows = grid.find_elements_by_tag_name("tr");
        row_count = len(grid_rows)
        if (row_count >= 2):
            Total_Count = driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_lblcount').text
            print("::" +Total_Count)
            sheetname['b58'] = "Matching - Mismatch Risk based Other :: Loading fine !"
            sheetname['a58'] = "Matching - Mismatch Risk based Other "
        elif (row_count == 1 ):
            error_sheet['b45'] = "Mismatch Risk based Others"
            error_sheet['c45'] = "No data found "
            error_sheet['b46'] = "Month::" + "" + Curr_Return_Month
            error_sheet['b47'] = "Grp range::" + "" + Other_range
            error_sheet['b48'] = " Year::" + "" + Curr_Year

    except Exception as e:
        print("The Exception found:", str(e))
        destination = "Media//MismatchRiskbased_other.png"
        driver.save_screenshot(destination)
        print("Error Screenshot saved as ::", destination)
        error_sheet['b13'] = "Mismatch Risk based Others"
        error_sheet['c45'] = "Error page occurs "
        error_sheet['b46'] = "Month::" + "" + Curr_Coll_Month
        error_sheet['b47'] = "Grp range::" + "" + Other_range
        error_sheet['b48'] = " Year::" + "" + Curr_Year
    finally:
        print("Done : Matching_Mismatch_Riskbased......")

    # B,2# Mismatch Exccess ITC Claim ............

    time.sleep(5)
    driver.get(url + "/Reports/rpt_R3BR2AComp.aspx")
    time.sleep(3)
    select = Select(driver.find_element(By.ID, 'ctl00_dd_finyear'))
    select.select_by_visible_text("2019-2020")
    time.sleep(7)
    driver.find_element(By.ID,'ctl00_id_btn_go').click()
    #WebDriverWait(driver, 30).until(EC.visibility_of((ele)))
    #if element.is_displayed() == visibility else False
    time.sleep(7)
    sheetname['a62'] = "MAtching_Excess ITC Claim "
    #sheetname['b62'] = "MAtching_Excess ITC Claim "
    sheetname['b63'] = "Year:: 2019-2020"
    time.sleep(5)
    try:
        time.sleep(7)
        element = driver.find_element(By.XPATH, "//table[@id='ctl00_ContentPlaceHolder1_grd_main']")
        WebDriverWait(driver, 30).until(EC.visibility_of((element)))
        grid = driver.find_element(By.XPATH, "//table[@id='ctl00_ContentPlaceHolder1_grd_main']")
        grid_rows = grid.find_elements_by_tag_name("tr");
        row_count = len(grid_rows)
        if (row_count >= 2):
            Total_Count = driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_lblcount').text
            print(" ::" + Total_Count)
            sheetname['b62'] = "Excess ITC Claim :: Loading fine !"
            print("Matching Excess ITC Claim page is loading fine !")
        elif (row_count == 1):
            error_sheet['b49'] = "Mismatch Excess ITC Claim"
            error_sheet['c49'] = "No data found "
            error_sheet['b50'] = "Year:: 2019-2020"

    except Exception as e:
        print("The Exception found:", str(e))
        destination = "Media//ExcessITC_Claim.png"
        driver.save_screenshot(destination)
        print("Error Screenshot saved as ::", destination)
        error_sheet['b49'] = "Mismatch Excess ITC Claim"
        error_sheet['c49'] = "Error page occurs "
        error_sheet['b50'] = "Year:: 2019-2020"

    finally:
        print("Done: Excess ITC Claim......")

    #driver.get(url + "/Reports/Rpt_AllTaxReport.aspx?txt_rpttype=7")

    # B,3#  ITC of GST Yearly ............
    driver.get(url + "/Reports/Rpt_AllTaxReport.aspx?txt_rpttype=7")
    time.sleep(3)
    Range = 'ctl00_ddl_tax_payer_range'
    Go = 'ctl00_ContentPlaceHolder1_Button1'
    driver.find_element(By.ID, Range).send_keys(Grp_range)
    time.sleep(5)
    driver.find_element(By.ID, Go).click()
    time.sleep(7)
    sheetname['a64'] = "Matching _ ITC OF GST Yearly"
    sheetname['b65'] = "Grp range::" + "" + Grp_range
    sheetname['b66'] = " Year::" + "" + Curr_Year

    try:
        time.sleep(7)
        element = driver.find_element(By.XPATH, "//table[@id='grd_main']")
        WebDriverWait(driver, 20).until(EC.visibility_of((element)))
        grid = driver.find_element(By.XPATH, "//table[@id='grd_main']")
        grid_rows = grid.find_elements_by_tag_name("tr");
        row_count = len(grid_rows)
        if (row_count >= 2):
            Total_Count = driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_lbl_msg').text
            print("::" +Total_Count)
            sheetname['b64'] = "ITC of GST Yearly :: Loading fine !"
        elif(row_count == 1):
            error_sheet['b51'] = "ITC of GST Yearly"
            error_sheet['c51'] = "No data found "
            error_sheet['b52'] = "Grp range::" + "" + Grp_range
            error_sheet['b53'] = " Year::" + "" + Curr_Year


    except Exception as e:
        print("The Exception found:", str(e))
        destination = "Media//ITCOfGSTYearly.png"
        driver.save_screenshot(destination)
        print("Error Screenshot saved as ::", destination)
        error_sheet['b51'] = "ITC of GST Yearly"
        error_sheet['c51'] = "Error page occurs "
        error_sheet['b52'] = "Grp range::" + "" + Grp_range
        error_sheet['b53'] = " Year::" + "" + Curr_Year
    finally:
        print("Done :Rpt_AllTaxReport_Matching ITC GST Yearly........")

    # C # Field offices - GSTO ...........................

    driver.get(url + "/Reports/R3BTaxCollReport.aspx")
    time.sleep(3)
    driver.find_element(By.ID, "ctl00_MainContent_ddl_rtprd_month").send_keys(Curr_Return_Month)
    driver.find_element(By.ID, "ctl00_MainContent_ddl_rtprd_year").send_keys(Curr_Year)
    time.sleep(3)
    element = driver.find_element(By.XPATH, "//input[@id='ctl00_MainContent_btnshow']")
    element.send_keys(Keys.ENTER)
    try:
        alert = driver.switch_to.alert
        alert.accept()
    except NoAlertPresentException:
        print("No alert found")

    time.sleep(5)
    sheetname['a67'] = "Matching-Field Offices_GSTO "
    sheetname['b68'] = "Month::" + "" + Curr_Coll_Month
    sheetname['b69'] = " Year::" + "" + Curr_Year
    time.sleep(7)
    try:
        element = driver.find_element(By.XPATH, "//table[@id='grd_dvo']")
        WebDriverWait(driver, 20).until(EC.visibility_of((element)))
        grid = driver.find_element(By.XPATH, "//table[@id='grd_dvo']")
        grid_rows = grid.find_elements_by_tag_name("tr");
        row_count = len(grid_rows)
        if (row_count>= 2):
            sheetname['b67'] = "Filed offices GSTO :: Loading fine !"
            print("Field offices GSTO page is loading fine !")
        elif (row_count == 1 ):
            error_sheet['b54'] = "Filed offices GSTO"
            error_sheet['c54'] = "No data found "
            error_sheet['b55'] = "Month::" + "" + Curr_Coll_Month
            error_sheet['b56'] = " Year::" + "" + Curr_Year

    except Exception as e:
        print("The Exception found:", str(e))
        destination ="Media//FieldOffices_GSTO.png"
        driver.save_screenshot(destination)
        print("Error Screenshot saved as ::", destination)
        error_sheet['b54'] = "Field offices GSTO"
        error_sheet['c54'] = "Error page occurs"
        error_sheet['b55'] = "Month::" + "" + Curr_Coll_Month
        error_sheet['b56'] = " Year::" + "" + Curr_Year

    finally:
        print("Done : R3BTaxCollReport_Field Offices GSTO.......")


    # ... Summary report - Payments .........
    time.sleep(7)
    To_date = "31/07/2020"
    frm_date = "01/07/2020"
    driver.get(url + "/Reports/Payment_summary.aspx")
    time.sleep(3)
    driver.find_element(By.XPATH, "//input[@id='ctl00_MainContent_txt_from_date']").clear()
    driver.find_element(By.XPATH, "//input[@id='ctl00_MainContent_txt_from_date']").send_keys(frm_date)
    driver.find_element(By.XPATH, "//input[@id='ctl00_MainContent_txt_to_date']").clear()
    driver.find_element(By.XPATH, "//input[@id='ctl00_MainContent_txt_to_date']").send_keys(To_date)
    time.sleep(5)
    element = driver.find_element(By.ID, 'ctl00_MainContent_btn_go')
    element.send_keys(Keys.ENTER)
    time.sleep(5)
    sheetname['a71'] = "Summary - Payments"
    sheetname['b72'] = "From_Month ::" + "" + frm_date
    sheetname['b73'] = " To_Month ::" + "" + To_date
    time.sleep(5)
    try:
        #grid = "//table[@id='ctl00_MainContent_GV_List']"

        element = driver.find_element(By.XPATH, "//table[@id='ctl00_MainContent_grd_div']")
        WebDriverWait(driver, 20).until(EC.visibility_of((element)))
        #grid = driver.find_element(By.XPATH, "//table[@id='ctl00_MainContent_GV_List']")
        grid = driver.find_element(By.XPATH, "//table[@id='ctl00_MainContent_grd_div']")
        grid_rows = grid.find_elements_by_tag_name("tr");
        row_count = len(grid_rows)
        if (row_count >= 2):
           sheetname['b71'] = "Summary-Payments :: Loading fine !"
           #sheetname['a68'] = "Summary-Payments "

        elif (row_count == 1 ):
           #sheetname['c68'] = "Summary-Payments :: No data found"
           error_sheet['b57'] = "Summary Payments"
           error_sheet['c57'] = "No data found "
           error_sheet['b58'] = "From_Month ::" + "" + frm_date
           error_sheet['b59'] = " To_Month ::" + "" + To_date

    except Exception as e :
        print("The Exception found:", str(e))
        destination = "Media//Summary_Payments.png"
        driver.save_screenshot(destination)
        print("Error Screenshot saved as ::", destination)
        #sheetname['c68'] = "Summary Payments :: Error page occurs ! "
        error_sheet['b57'] = "Summary Payments"
        error_sheet['c57'] = "Error Page occurs !"
        error_sheet['b58'] = "From_Month ::" + "" + frm_date
        error_sheet['b59'] = " To_Month ::" + "" + To_date

    finally:
        print("Done : Payment_summary......")

    # ... Summary Report - Annual return #
    #.. GSTR9
    driver.get(url + "/Reports/rpt_gstr9.aspx?type=S")
    time.sleep(7)
    element = driver.find_element(By.XPATH, "//input[@id='ctl00_MainContent_btn_go']").click()
    # element.send_keys(Keys.ENTER)
    try:
        time.sleep(7)
        sheetname['a145'] = "Summary - Annual return - GSTR9 !"
        sheetname['b146'] = " Year::" + "" + Curr_Year
        grid_main = driver.find_element(By.XPATH, "//table[@id='grd_dgsto']")
        WebDriverWait(driver, 20).until(EC.visibility_of((element)))
        grid_main = driver.find_element(By.XPATH, "//table[@id='grd_dgsto']")
        grid_rows = grid_main.find_elements_by_tag_name("tr")
        row_count = len(grid_rows)
        if (row_count >= 2):
            sheetname['b145'] = "Summary-AnnualReturn-GSTR9 :: Loading fine !"
            #sheetname['a146'] = "Summary-AnnualReturn-GSTR9"
            print("Summary-AnnualReturn-GSTR9 :: Loading fine !")
        elif (row_count == 1 ):
            error_sheet['b60'] = "Summary- AnnualReturn-GSTR9"
            error_sheet['c60'] = "No data found "

    except Exception as e:
        print(e)
        destination = "Media//Summary-AnnualReturn-GSTR9.png"
        driver.save_screenshot(destination)
        print("Error Screenshot saved as ::", destination)
        error_sheet['b60'] = "Summary-AnnualReturn-GSTR9"
        error_sheet['c60'] = "No Such Element/Stale Element"

    # GSTR9A...
    time.sleep(5)
    driver.get(url + "/Reports/rpt_gstr9.aspx?type=S")
    try:
        time.sleep(7)
        driver.find_element(By.XPATH, "//input[@id='ctl00_MainContent_rdReturnType_1']").click()
        time.sleep(7)
        element = driver.find_element(By.XPATH, "//input[@id='ctl00_MainContent_btn_go']")
        element.send_keys(Keys.ENTER)
        time.sleep(5)
        sheetname['a147'] = "Summary - Annual return - GSTR9A !"
        sheetname['b148'] = " Year::" + "" + Curr_Year
        grid_main = driver.find_element(By.XPATH, "//table[@id='grd_dgsto']")
        WebDriverWait(driver, 20).until(EC.visibility_of((element)))
        grid_main = driver.find_element(By.XPATH, "//table[@id='grd_dgsto']")
        grid_rows = grid_main.find_elements_by_tag_name("tr")
        row_count = len(grid_rows)
        if (row_count >= 2):
            sheetname['b147'] = "Summary-Annual return-GSTR9A :: Loading fine!"
            sheetname['a147'] = "Summary-Annual return-GSTR9A "
            print("Summary - Annual return - GSTR9A Page is loading fine !")
        elif (row_count == 1 ):
            error_sheet['b61'] = "SummaryAnnualReturn-GSTR9A"
            error_sheet['c61'] = "No data found"

    except Exception as e:
        print("The Exception found:", str(e))
        destination = "Media//SummaryAnnualReturn-GSTR9A.png"
        driver.save_screenshot(destination)
        print("Error Screenshot saved as ::", destination)
        error_sheet['b61'] = "Summary-AnnualReturn-GSTR9A"
        error_sheet['c61'] = "No Such Element/ Stale Element"

    finally:
        print("Done:Summary-Annual return GSTR9A......")

    # 5 # MIS Registration Statctics ...
    driver.get(url + "/Reports/rptRegistrationStat.aspx")
    driver.find_element(By.ID, 'ctl00_ddl_Month').send_keys(Curr_Return_Month)
    driver.find_element(By.ID, 'ctl00_ddl_year').send_keys(Curr_Year)
    Go=driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_btn_go')
    Go.send_keys(Keys.ENTER)
    time.sleep(7)
    sheetname['a74'] = "MiS Report_Registration Statictics "
    sheetname['b75'] = "Month::" + "" + Curr_Return_Month
    sheetname['b76'] = " Year::" + "" + Curr_Year
    try:
        time.sleep(7)
        element = driver.find_element(By.XPATH, "//table[@id='grd_main']")
        WebDriverWait(driver, 20).until(EC.visibility_of((element)))
        grid = driver.find_element(By.XPATH, "//table[@id='grd_main']")
        grid_rows = grid.find_elements_by_tag_name("tr");
        row_count = len(grid_rows)
        if (row_count >= 2):
            sheetname['b74'] = "MiS Report_Registration Statictics :: Loading fine !"
            #sheetname['a76'] = "MiS Report_Registration Statictics "
        elif (row_count == 1 ):
            error_sheet['b61'] = "MiS Registration Statictics"
            error_sheet['c61'] = "No data found"
            error_sheet['b62'] = "Month::" + "" + Curr_Return_Month
            error_sheet['b63'] = " Year::" + "" + Curr_Year

    except Exception as e:
        print("The Exception found:", str(e))
        destination = " Media//MiSRegistration_Statictics.png"
        driver.save_screenshot(destination)
        print("Error Screenshot saved as ::", destination)
        error_sheet['b61'] = "MiS Registration Statictics"
        error_sheet['c61'] = "Error page occurs"

        error_sheet['b62'] = "Month::" + "" + Curr_Return_Month
        error_sheet['b63'] = " Year::" + "" + Curr_Year

    finally:
        print("DOne: Mis Registration Statictics......")

    # 5 # MIS Report Final return ............
    time.sleep(5)
    driver.get(url + "/Reports/rpt_GSTR10.aspx")
    frm_date = "1-06-2020"
    to_date = "30-06-2020"
    time.sleep(5)
    driver.find_element(By.ID, 'ctl00_MainContent_txt_fromdate').clear()
    time.sleep(3)
    driver.find_element(By.ID, 'ctl00_MainContent_txt_fromdate').send_keys(frm_date)
    driver.find_element(By.ID, 'ctl00_MainContent_txt_todate').clear()
    driver.find_element(By.ID, 'ctl00_MainContent_txt_todate').send_keys(to_date)
    time.sleep(5)
    driver.find_element(By.ID, 'ctl00_MainContent_btn_go').click()
    sheetname['a77'] = "MiS_Report_Final_return "
    sheetname['b78'] = "Month::" + "" + frm_date
    sheetname['b79'] = " Year::" + "" + to_date
    time.sleep(5)
    try:
        #grid = "//table[@id='grdZone']"
        grid = "//table[@id='grd_dgsto']"
        element = driver.find_element(By.XPATH, grid)
        WebDriverWait(driver, 30).until(EC.visibility_of((element)))
        grid = driver.find_element(By.XPATH, "//table[@id='grd_dgsto']")
        grid_rows = grid.find_elements_by_tag_name("tr");
        row_count = len(grid_rows)
        if (row_count >= 2):
            sheetname['b79'] = "MiS Report_Final Return :: Loading fine !"
        elif (row_count == 1 ):
            error_sheet['b64'] = "MiS Report_Final Return"
            error_sheet['c64'] = "No data found"
            error_sheet['b65'] = "Month::" + "" + frm_date
            error_sheet['b66'] = " Year::" + "" + to_date
    except Exception as e:
        print("The Exception found:", str(e))
        destination = "Media//MiSFinal_Return.png"
        driver.save_screenshot(destination)
        print("Error Screenshot saved as ::", destination)
        #sheetname['c21'] = "MiS Report_Final Return :: Error page occurs ! "
        error_sheet['b64'] = "MiS_Final _ Return"
        error_sheet['c64'] = "Error page occurs"
        error_sheet['b65'] = "Month::" + "" + frm_date
        error_sheet['b66'] = " Year::" + "" + to_date

    finally:
        print("Done:MIS Report Final Return......")


    # Non Filers ..........
    # Non Filings - Analytic_NonFilers_SGST.......

    driver.get(url + "/Reports/rptr3btopnonfilers.aspx")

    Month = "//*[@id='ctl00_ddl_Month']"
    Range = "//*[@id='ctl00_ddl_tax_payer_range']"
    Year = "ctl00_ddl_year"
    Go = "//input[@id='ctl00_btn_go']"
    #####..............................................
    month = driver.find_element(By.XPATH, Month).send_keys(Curr_Return_Month)
    range = driver.find_element(By.XPATH, Range).send_keys(Grp_range)
    year = driver.find_element(By.ID, Year).send_keys(Curr_Year)
    sheetname['a80'] = "Analytic_NonFilers_SGST_Top"
    sheetname['b81'] = "Month::" + "" + Curr_Return_Month
    sheetname['b82'] = "Grp range::" + "" + Grp_range
    sheetname['b83'] = " Year::" + "" + Curr_Year
    try:
        time.sleep(7)
        driver.find_element(By.XPATH, "//input[@id='ctl00_btn_go']").click()
        #go.send_keys(Keys.ENTER)
        time.sleep(5)
        page_title = driver.title
        element = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_grd_main"]')
        WebDriverWait(driver, 20).until(EC.visibility_of((element)))
        grd_main = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_grd_main"]')
        grid_rows = grd_main.find_elements_by_tag_name("tr");
        row_count = len(grid_rows)
        if  (row_count >= 2):
             Total_Count = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_lblcount").text
             print("::" +Total_Count)
             sheetname['b80'] = "R3BNonFilers_SGST_Top :: Loading fine !"
             #sheetname['a83'] = "R3BNonFilers_SGST_Top "
             print("R3B_NonFilers_SGST_Top- page is loading fine")
        elif (row_count == 1 ):
            #heetname['c83'] = "R3BNonFilers_SGST_Top :: No data found !"
            error_sheet['b67'] = "R3BNonFilers_SGST_Top"
            error_sheet['c67'] = "No data found"
            error_sheet['b68'] = "Month::" + "" + Curr_Return_Month
            error_sheet['b69'] = "Grp range::" + "" + Grp_range
            error_sheet['b70'] = " Year::" + "" + Curr_Year

    except Exception as e:
        print("The Exception found:", str(e))
        destination = "Media//R3BNonFilersSGST_Top.png"
        driver.save_screenshot(destination)
        print("Error Screenshot saved as ::", destination)

        error_sheet['b67'] = "R3BNonFilers_SGST_Top"

        error_sheet['c67'] = "Error page occurs"
        error_sheet['b68'] = "Month::" + "" + Curr_Return_Month
        error_sheet['b69'] = "Grp range::" + "" + Grp_range
        error_sheet['b70'] = " Year::" + "" + Curr_Year

    # Non filers -SGST-middle
    driver.get(url + "/Reports/rptr3btopnonfilers.aspx")
    driver.find_element(By.XPATH, "//input[@id='ctl00_rdb_taxpayer_type_1']").click()
    month = driver.find_element(By.XPATH, Month).send_keys(Curr_Return_Month)
    range = driver.find_element(By.XPATH, Range).send_keys(Mid_range)
    year = driver.find_element(By.ID, Year).send_keys(Curr_Year)
    driver.find_element(By.XPATH, Go).click()
    time.sleep(5)
    page_title = driver.title
    sheetname['a84'] = "R3BNonFilers_SGST_Middle "
    sheetname['b85'] = "Month::" + "" + Curr_Return_Month
    sheetname['b86'] = "Grp range::" + "" + Mid_range
    sheetname['b87'] = " Year::" + "" + Curr_Year
    try:
        element = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_grd_main"]')
        WebDriverWait(driver, 20).until(EC.visibility_of((element)))
        grd_main = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_grd_main"]')
        grid_rows = grd_main.find_elements_by_tag_name("tr");
        row_count = len(grid_rows)
        if (row_count >= 2):
            Total_Count = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_lblcount").text
            print("::" +Total_Count)
            sheetname['b84'] = "R3BNonFilers_SGST_Middle :: Loading fine !"

            print("R3B_NonFilers_SGST_Middle - page is loading fine")
        elif (row_count == 1 ):
            error_sheet['b71'] = "R3BNonFilers_SGST_Middle"
            error_sheet['c71'] = "No data found"
            error_sheet['b72'] = "Month::" + "" + Curr_Return_Month
            error_sheet['b73'] = "Grp range::" + "" + Mid_range
            error_sheet['b74'] = " Year::" + "" + Curr_Year
    except Exception as e:
        print("The Exception found:", str(e))
        destination = "Media//R3BNonFilersSGST_Middle.png"
        driver.save_screenshot(destination)
        print("Error Screenshot saved as ::", destination)
        error_sheet['b71'] = "R3BNonFilers_SGST_Middle"
        error_sheet['c71'] = "Error page occurs"
        error_sheet['b72'] = "Month::" + "" + Curr_Return_Month
        error_sheet['b73'] = "Grp range::" + "" + Mid_range
        error_sheet['b74'] = " Year::" + "" + Curr_Year
    # Non filers -SGST-other...
    driver.get(url + "/Reports/rptr3btopnonfilers.aspx")
    driver.find_element(By.XPATH, "//input[@id='ctl00_rdb_taxpayer_type_2']").click()
    month = driver.find_element(By.XPATH, Month).send_keys(Curr_Return_Month)
    range = driver.find_element(By.XPATH, Range).send_keys(Other_range)
    year = driver.find_element(By.ID, Year).send_keys(Curr_Year)
    driver.find_element(By.XPATH, Go).click()
    time.sleep(5)
    page_title = driver.title
    sheetname['b89'] = "Month::" + "" + Curr_Return_Month
    sheetname['b90'] = "Grp range::" + "" + Other_range
    sheetname['b91'] = " Year::" + "" + Curr_Year
    try:
        element = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_grd_main"]')
        WebDriverWait(driver, 20).until(EC.visibility_of((element)))

        sheetname['a88'] = " Non Filers SGST other"
        Total_Count = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_lblcount").text
        print("::" +Total_Count)
        grd_main = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_grd_main"]')
        grid_rows = grd_main.find_elements_by_tag_name("tr");
        row_count = len(grid_rows)
        if(row_count>=2):
            sheetname['b88'] = " Non Filers SGST other :: Loading fine !"
        elif (row_count == 1):
            error_sheet['b71'] = "R3BNonFilers_SGST_Other"
            error_sheet['c71'] = "No data found"
            error_sheet['b72'] = "Month::" + "" + Curr_Return_Month
            error_sheet['b73'] = "Grp range::" + "" + Other_range
            error_sheet['b74'] = " Year::" + "" + Curr_Year
    except Exception as e:
        print("The Exception found:", str(e))
        destination = "Media//R3BNonFilersSGST_Others.png"
        driver.save_screenshot(destination)
        print("Error Screenshot saved as ::", destination)
        print("The Exception found:", str(e))
        error_sheet['b71'] = "R3BNonFilers_SGST_Other"
        error_sheet['c71'] = "Error page occurs"
        error_sheet['b72'] = "Month::" + "" + Curr_Return_Month
        error_sheet['b73'] = "Grp range::" + "" + Other_range
        error_sheet['b74'] = " Year::" + "" + Curr_Year
    finally:
        print("Done : rptr3btopnonfilers_Nonfilings  SGST......")

   # Nil filers continuously 3 months .............#
    driver.get(url + "/Reports/rpt_r3b_nil_filers.aspx")
    time.sleep(5)
    month = driver.find_element(By.XPATH, "//*[@id='ctl00_ddl_Month']").send_keys(Curr_Return_Month)
    range = driver.find_element(By.XPATH, "//*[@id='ctl00_ddl_tax_payer_range']").send_keys(Grp_range)
    year = driver.find_element(By.ID, "ctl00_ddl_year").send_keys(Curr_Year)
    driver.find_element(By.XPATH, "//input[@id='ctl00_btn_go']").click()
    time.sleep(5)
    page_title = driver.title
    print(page_title)
    sheetname['a92'] = "Nil FIlers continuously 3 months_Top "
    sheetname['b93'] = "Month::" + "" + Curr_Return_Month
    sheetname['b94'] = "Grp range::" + "" + Grp_range
    sheetname['b95'] = " Year::" + "" + Curr_Year
    try:
        element = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_grd_main"]')
        WebDriverWait(driver, 20).until(EC.visibility_of((element)))
        grd_main = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_grd_main"]')
        grid_rows = grd_main.find_elements_by_tag_name("tr");
        row_count = len(grid_rows)
        if (row_count >= 2):
            Total_Count = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_lblcount").text
            print("::" +Total_Count)
            print("NilFilers cont.3 months Top :: Loading fine")
            sheetname['b92'] = "NilFilers cont.3 months Top :: Loading fine !"
        elif (row_count == 1 ):
           # sheetname['c95'] = "NilFilers cont.3 months Top :: No data found"
            error_sheet['b75'] = "NilFilers cont.3 months Top"
            error_sheet['c75'] = "No data found"
            error_sheet['b76'] = "Month::" + "" + Curr_Return_Month
            error_sheet['b77'] = "Grp range::" + "" + Grp_range
            error_sheet['b78'] = " Year::" + "" + Curr_Year
    except Exception as e:
        print("The Exception found:", str(e))
        destination = "Media//NilFilers-cont.3months_Top.png"
        driver.save_screenshot(destination)
        print("Error Screenshot saved as ::", destination)
        error_sheet['c75'] = "Error page occurs"
        error_sheet['b75'] = "NilFilers cont.3 months Top"
        error_sheet['b76'] = "Month::" + "" + Curr_Return_Month
        error_sheet['b77'] = "Grp range::" + "" + Grp_range
        error_sheet['b78'] = " Year::" + "" + Curr_Year
    # Middle ...
    driver.get(url + "/Reports/rpt_r3b_nil_filers.aspx")
    time.sleep(5)
    month = driver.find_element(By.XPATH, "//*[@id='ctl00_ddl_Month']").send_keys(Curr_Return_Month)
    range = driver.find_element(By.XPATH, "//*[@id='ctl00_ddl_tax_payer_range']").send_keys(Mid_range)
    year = driver.find_element(By.ID, "ctl00_ddl_year").send_keys(Curr_Year)
    driver.find_element(By.XPATH, "//input[@id='ctl00_btn_go']").click()
    time.sleep(5)
    page_title = driver.title
    print(page_title)
    sheetname['b97'] = "Month::" + "" + Curr_Return_Month
    sheetname['b98'] = "Grp range::" + "" + Mid_range
    sheetname['b99'] = " Year::" + "" + Curr_Year
    sheetname['a96'] = "NilFIlers cont.3MonthsMiddle "
    try:
        element = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_grd_main"]')
        WebDriverWait(driver, 20).until(EC.visibility_of((element)))
        grd_main = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_grd_main"]')
        grid_rows = grd_main.find_elements_by_tag_name("tr");
        row_count = len(grid_rows)
        if (row_count >= 2):
            Total_Count = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_lblcount").text
            print("::" +Total_Count)
            sheetname['b96'] = "NilFIlers cont.3MonthsMiddle :: Loading fine !"
        elif (row_count == 1 ):
            error_sheet['b79'] = "NilFIlers cont.3MonthsMiddle"
            error_sheet['c79'] = "No data found"
            error_sheet['b80'] = "Month::" + "" + Curr_Return_Month
            error_sheet['b81'] = "Grp range::" + "" + Mid_range
            error_sheet['b82'] = " Year::" + "" + Curr_Year

    except Exception as e:
        print("The Exception found:", str(e))
        destination = "Media//NilFIlerscont.3Months_Middle.png"
        driver.save_screenshot(destination)
        print("Error Screenshot saved as ::", destination)
        error_sheet['b79'] = "NilFIlers cont.3MonthsMiddle"
        error_sheet['c79'] = "Error page occurs"
        error_sheet['b80'] = "Month::" + "" + Curr_Return_Month
        error_sheet['b81'] = "Grp range::" + "" + Mid_range
        error_sheet['b82'] = " Year::" + "" + Curr_Year

    # Others.....
    driver.get(url + "/Reports/rpt_r3b_nil_filers.aspx")
    time.sleep(5)
    driver.find_element(By.XPATH, "//input[@id='ctl00_rdb_taxpayer_type_2']").click()
    month = driver.find_element(By.XPATH, "//*[@id='ctl00_ddl_Month']").send_keys(Curr_Return_Month)
    range = driver.find_element(By.XPATH, "//*[@id='ctl00_ddl_tax_payer_range']").send_keys(Other_range)
    year = driver.find_element(By.ID, "ctl00_ddl_year").send_keys(Curr_Year)
    driver.find_element(By.XPATH, "//input[@id='ctl00_btn_go']").click()
    time.sleep(5)
    page_title = driver.title
    print(page_title)
    sheetname['b101'] = "Month::" + "" + Curr_Return_Month
    sheetname['b102'] = "Grp range::" + "" + Other_range
    sheetname['b103'] = " Year::" + "" + Curr_Year
    sheetname['a100'] = "NilFIlers cont.3MonthsOthers"
    try:
        element = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_grd_main"]')
        WebDriverWait(driver, 20).until(EC.visibility_of((element)))
        grd_main = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_grd_main"]')
        grid_rows = grd_main.find_elements_by_tag_name("tr");
        row_count = len(grid_rows)
        if (row_count >= 2):
            Total_Count = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_lblcount").text
            print("::" +Total_Count)
            print("Non FIlers continuously 3 months_others - page is loading fine")
            sheetname['b100'] = "NilFIlers cont.3MonthsOthers :: Loading fine !"

        elif (row_count == 1 ):
            error_sheet['b83'] = "NilFIlers cont.3MonthsOthers"
            error_sheet['c83'] = "No data found"
            error_sheet['b84'] = "Month::" + "" + Curr_Return_Month
            error_sheet['b85'] = "Grp range::" + "" + Other_range
            error_sheet['b86'] = " Year::" + "" + Curr_Year

    except Exception as e:
        print("The Exception found:", str(e))
        destination = "Media//NilFIlerscont.3Months_Others .png"
        driver.save_screenshot(destination)
        print("Error Screenshot saved as ::", destination)
        error_sheet['c83'] = "Error page occurs"
        error_sheet['b83'] = "NilFIlers cont.3MonthsOthers"
        error_sheet['b84'] = "Month::" + "" + Curr_Return_Month
        error_sheet['b85'] = "Grp range::" + "" + Other_range
        error_sheet['b86'] = " Year::" + "" + Curr_Year
    finally:
        print("Done : rpt_r3b_nil_filers_R3B Nil FIlers......")

    # .... TDS/TCS.............................
    driver.get(url + "/Reports/Rpt_TDSNonfilers.aspx")
    month = driver.find_element(By.XPATH, "//*[@id='ctl00_ddl_Month']").send_keys(Curr_Return_Month)
    year = driver.find_element(By.ID, "ctl00_ddl_year").send_keys(Curr_Year)
    driver.find_element(By.XPATH, "//input[@id='ctl00_ContentPlaceHolder1_Button']").click()
    time.sleep(5)
    page_title = driver.title
    print(page_title)
    sheetname['a104'] = " TDS/TCS Important "
    sheetname['b106'] = "Month::" + "" + Curr_Return_Month
    sheetname['b107'] = " Year::" + "" + Curr_Year
    sheetname['a105'] = "TDS Important"
    try:
        element = driver.find_element(By.XPATH, '//*[@id="grd_main"]')
        WebDriverWait(driver, 20).until(EC.visibility_of((element)))
        Total_Count = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_lblrecordsno").text
        print("::" +Total_Count)
        grd_main = driver.find_element(By.XPATH, '//*[@id="grd_main"]')
        grid_rows = grd_main.find_elements_by_tag_name("tr");
        row_count = len(grid_rows)
        if (row_count >= 2):
            sheetname['b105'] = "TDS Important :: Loading fine !"
            #print("TDS Important- Page is loading fine !")
        elif (row_count == 1):
            error_sheet['b87'] = "TDS Important"
            error_sheet['c87'] = "No data found"
            error_sheet['b88'] = "Month::" + "" + Curr_Return_Month
            error_sheet['b89'] = " Year::" + "" + Curr_Year

    except Exception as e:
        print("The Exception found:", str(e))
        destination = "Media//TDS-Important.png"
        driver.save_screenshot(destination)
        print("Error Screenshot saved as ::", destination)
        error_sheet['b87'] = "TDS Important"

        error_sheet['c87'] = "Error page occurs"
        error_sheet['b88'] = "Month::" + "" + Curr_Return_Month
        error_sheet['b89'] = " Year::" + "" + Curr_Year

    # TCS Important .............................
    driver.get(url + "/Reports/Rpt_TDSNonfilers.aspx")
    time.sleep(5)
    driver.find_element(By.XPATH,"//input[@id='ctl00_rdb_taxpayer_tdstcstype_1']").click()
    time.sleep(3)
    month = driver.find_element(By.XPATH, "//*[@id='ctl00_ddl_Month']").send_keys(Curr_Return_Month)
    year = driver.find_element(By.ID, "ctl00_ddl_year").send_keys(Curr_Year)
    driver.find_element(By.XPATH, "//input[@id='ctl00_ContentPlaceHolder1_Button']").click()
    time.sleep(5)
    page_title = driver.title
    print(page_title)
    sheetname['b109'] = "Month::" + "" + Curr_Return_Month
    sheetname['b110'] = " Year::" + "" + Curr_Year
    sheetname['a108'] = "TCS Important"
    try:
        element = driver.find_element(By.XPATH, "//table[@id='grd_main']")
        WebDriverWait(driver, 20).until(EC.visibility_of((element)))
        Total_Count = driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_lblrecordsno').text
        print("::" +Total_Count)
        grd_main = driver.find_element(By.XPATH, "//table[@id='grd_main']")
        grid_rows = grd_main.find_elements_by_tag_name("tr");
        row_count = len(grid_rows)
        if (row_count >= 2):
            sheetname['b108'] = "TCS Important :: Loading fine !"
        elif (row_count == 1 ):
            error_sheet['b90'] = "TCS Important"
            error_sheet['c90'] = "Error page occurs"
            error_sheet['b91'] = "Month::" + "" + Curr_Return_Month
            error_sheet['b92'] = " Year::" + "" + Curr_Year

    except Exception as e:
        print("The Exception found:", str(e))
        print("Error page occurs")
        destination = "Media//TCS-Important.png"
        driver.save_screenshot(destination)
        print("Error Screenshot saved as ::", destination)
        error_sheet['b90'] = "TCS Important"
        error_sheet['c90'] = "Error page occurs"
        error_sheet['b91'] = "Month::" + "" + Curr_Return_Month
        error_sheet['b92'] = " Year::" + "" + Curr_Year
    finally:
        print("Done: TDS / TCS important...... ")


    # ... Non Filers Risk based ...# ... Non Filers Risk based ...
    driver.get(url + "/Reports/rptR1R2AR7R8EWB_NF.aspx")
    time.sleep(2)
    month = driver.find_element(By.XPATH, "//*[@id='ctl00_ddl_Month']").send_keys(Curr_Return_Month)
    range = driver.find_element(By.XPATH, "//*[@id='ctl00_ddl_tax_payer_range']").send_keys(Grp_range)
    year = driver.find_element(By.ID, "ctl00_ddl_year").send_keys(Curr_Year)
    driver.find_element(By.XPATH, "//input[@id='ctl00_btn_go']").click()
    time.sleep(5)
    page_title = driver.title
    print(page_title)
    #sheetname['a110'] = "NonFilers Risk based "
    sheetname['b114'] = "Month::" + "" + Curr_Return_Month
    sheetname['b115'] = " Year::" + "" + Curr_Year
    sheetname['b116'] = "range::" + "" + Grp_range
    try:
        element = driver.find_element(By.XPATH, "//*[@id='grd_main']")
        WebDriverWait(driver, 20).until(EC.visibility_of((element)))
        grd_main = driver.find_element(By.XPATH, "//*[@id='grd_main']")
        grid_rows = grd_main.find_elements_by_tag_name("tr");
        row_count = len(grid_rows)
        if (row_count >= 2):
            sheetname['b113'] = "NonFilersRiskBasedTop :: Loading fine !"
            sheetname['a113'] = "NonFilersRiskBasedTop"
            Total_Count = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_lblcount").text
            print(Total_Count)
            print("NonFilers Risk basedTop - Page is loading fine !")
        elif (row_count == 1 ):
            error_sheet['b93'] = "NonFilersRiskBasedTop"
            error_sheet['c93'] = "No data found"
            error_sheet['b94'] = "Month::" + "" + Curr_Return_Month
            error_sheet['b95'] = " Year::" + "" + Curr_Year
            error_sheet['b96'] = "range::" + "" + Grp_range
    except Exception as e:
        print("The Exception found:", str(e))
        destination = "Media//NonFilersRiskBasedTop.png"
        driver.save_screenshot(destination)
        print("Error Screenshot saved as ::", destination)
        error_sheet['b93'] = "NonFilersRiskBasedTop"
        error_sheet['c93'] = "Error page occurs "
        error_sheet['b94'] = "Month::" + "" + Curr_Return_Month
        error_sheet['b95'] = " Year::" + "" + Curr_Year
        error_sheet['b96'] = "range::" + "" + Grp_range
    # middle .........
    driver.get(url + "/Reports/rptR1R2AR7R8EWB_NF.aspx")
    time.sleep(4)
    driver.find_element(By.XPATH, "//input[@id= 'ctl00_rdb_taxpayer_type_1']").click()
    month = driver.find_element(By.XPATH, "//*[@id='ctl00_ddl_Month']").send_keys(Curr_Return_Month)
    range = driver.find_element(By.XPATH, "//*[@id='ctl00_ddl_tax_payer_range']").send_keys(Mid_range)
    year = driver.find_element(By.ID, "ctl00_ddl_year").send_keys(Curr_Year)
    driver.find_element(By.XPATH, "//input[@id='ctl00_btn_go']").click()
    time.sleep(5)
    page_title = driver.title
    print(page_title)
    sheetname['b118'] = "Month::" + "" + Curr_Return_Month
    sheetname['b119'] = " Year::" + "" + Curr_Year
    sheetname['b120'] = "range::" + "" + Mid_range
    sheetname['a117'] = "NonFilersRiskBasedMiddle"
    try:
        element = driver.find_element(By.XPATH, "//*[@id='grd_main']")
        WebDriverWait(driver, 20).until(EC.visibility_of((element)))
        grd_main = driver.find_element(By.XPATH, "//*[@id='grd_main']")
        grid_rows = grd_main.find_elements_by_tag_name("tr");
        row_count = len(grid_rows)
        if (row_count >= 2):
            Total_Count = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_lblcount").text
            print(Total_Count)
            sheetname['b117'] = "NonFilersRiskBasedMiddle :: Loading fine !"
            print("NonFilers Risk basedMiddle - Page is loading fine !")
        elif (row_count == 1 ):
            error_sheet['b97'] = "NonFilersRiskBasedMiddle"
            error_sheet['c97'] = "No data found"
            error_sheet['b98'] = "Month::" + "" + Curr_Return_Month
            error_sheet['b99'] = " Year::" + "" + Curr_Year
            error_sheet['b100'] = "range::" + "" + Mid_range
    except Exception as e:
        print("The Exception found:", str(e))
        destination = "Media//NonFilersRiskBasedMiddle.png"
        driver.save_screenshot(destination)
        print("Error Screenshot saved as ::", destination)
        error_sheet['b97'] = "NonFilersRiskBasedMiddle"
        error_sheet['c97'] = "Error page occurs"
        error_sheet['b98'] = "Month::" + "" + Curr_Return_Month
        error_sheet['b99'] = " Year::" + "" + Curr_Year
        error_sheet['b100'] = "range::" + "" + Mid_range
    # Others .............
    driver.get(url + "/Reports/rptR1R2AR7R8EWB_NF.aspx")
    time.sleep(2)
    driver.find_element(By.XPATH, "//input[@id= 'ctl00_rdb_taxpayer_type_2']").click()
    month = driver.find_element(By.XPATH, "//*[@id='ctl00_ddl_Month']").send_keys(Curr_Return_Month)
    range = driver.find_element(By.XPATH, "//*[@id='ctl00_ddl_tax_payer_range']").send_keys(Other_range)
    year = driver.find_element(By.ID, "ctl00_ddl_year").send_keys(Curr_Year)
    driver.find_element(By.XPATH, "//input[@id='ctl00_btn_go']").click()
    time.sleep(5)
    page_title = driver.title
    print(page_title)
    time.sleep(2)
    sheetname['a121'] = "NonFilersRiskBasedOthers"
    sheetname['b122'] = "Month::" + "" + Curr_Return_Month
    sheetname['b123'] = "Year::" + "" + Curr_Year
    sheetname['b124'] = "range::" + "" + Other_range

    try:
        element = driver.find_element(By.XPATH, "//*[@id='grd_main']")
        WebDriverWait(driver, 20).until(EC.visibility_of((element)))
        grd_main = driver.find_element(By.XPATH, "//*[@id='grd_main']")
        grid_rows = grd_main.find_elements_by_tag_name("tr");
        row_count = len(grid_rows)
        if (row_count >= 2):
            Total_Count = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_lblcount").text
            print(Total_Count)
            sheetname['b121'] = "NonFilersRiskBasedOthers :: Loading fine !"

            print("NonFilers Risk based other- Page is loading fine !")
        elif (row_count == 1 ):
            error_sheet['b101'] = "NonFilersRiskBasedOthers"
            error_sheet['c101'] = "No data found"
            error_sheet['b102'] = "Month::" + "" + Curr_Return_Month
            error_sheet['b103'] = " Year::" + "" + Curr_Year
            error_sheet['b104'] = "range::" + "" + Other_range
    except Exception as e:
            print("The Exception found:", str(e))
            destination = "Media//NonFilersRiskBasedOthers.png"
            driver.save_screenshot(destination)
            print("Error Screenshot saved as ::", destination)
            error_sheet['b101'] = "NonFilersRiskBasedOthers"
            error_sheet['c101'] = "Error page occurs"
            error_sheet['b102'] = "Month::" + "" + Curr_Return_Month
            error_sheet['b103'] = " Year::" + "" + Curr_Year
            error_sheet['b104'] = "range::" + "" + Other_range
    finally:
        print("Done ::Non FIlers Risk based...... ")

    # Non Filers - Total GST...

    driver.get(url + "/Reports/rptR3BTopNonFilers_Prev.aspx")
    time.sleep(2)
    month = driver.find_element(By.XPATH, "//*[@id='ctl00_ddl_Month']").send_keys(Curr_Return_Month)
    time.sleep(3)
    range = driver.find_element(By.XPATH, "//*[@id='ctl00_ddl_tax_payer_range']").send_keys(Grp_range)
    time.sleep(3)
    year = driver.find_element(By.ID, "ctl00_ddl_year").send_keys(Curr_Year)
    time.sleep(3)
    sheetname['a125'] = "NonFilers Total GST "
    sheetname['b126'] = "Month::" + "" + Curr_Return_Month
    sheetname['b127'] = " Year::" + "" + Curr_Year
    sheetname['b128'] = "range::" + "" + Grp_range

    try:
        time.sleep(7)
        element = driver.find_element(By.XPATH, "//input[@id='ctl00_btn_go']")
        element.send_keys(Keys.ENTER)
        time.sleep(5)
        page_title = driver.title
        print(page_title)
        element = driver.find_element(By.XPATH,"//table[@id='ctl00_ContentPlaceHolder1_grd_main']")
        WebDriverWait(driver, 30).until(EC.visibility_of((element)))
        grd_main = driver.find_element(By.XPATH,"//table[@id='ctl00_ContentPlaceHolder1_grd_main']")
        grid_rows = grd_main.find_elements_by_tag_name("tr");
        row_count = len(grid_rows)
        if (row_count >= 2):
            Total_Count = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_lblcount").text
            print("::" +Total_Count)
            sheetname['b125'] = "NonFilersTotalGSTTop :: Loading fine !"
            sheetname['a125'] = "NonFilersTotalGSTTop"
            print("NonFilers Total GST Top Page is loading fine !")
        elif (row_count == 1 ):
            error_sheet['b105'] = "NonFilersTotalGSTTop"
            error_sheet['c105'] = "No data found"
            error_sheet['b106'] = "Month::" + "" + Curr_Return_Month
            error_sheet['b107'] = " Year::" + "" + Curr_Year
            error_sheet['b108'] = "range::" + "" + Grp_range
    except Exception as e:
        print("The Exception found:", str(e))
        destination = "Media//NonFilersTotalGSTTop.png"
        driver.save_screenshot(destination)
        print("Error Screenshot saved as ::", destination)
        error_sheet['b105'] = "NonFilersTotalGSTTop"
        error_sheet['c105'] = "Error page occurs"
        error_sheet['b106'] = "Month::" + "" + Curr_Return_Month
        error_sheet['b107'] = " Year::" + "" + Curr_Year
        error_sheet['b108'] = "range::" + "" + Grp_range
    # middle .........................................
    driver.get(url + "/Reports/rptR3BTopNonFilers_Prev.aspx")
    time.sleep(2)
    driver.find_element(By.XPATH, "//input[@id= 'ctl00_rdb_taxpayer_type_1']").click()
    month = driver.find_element(By.XPATH, "//*[@id='ctl00_ddl_Month']").send_keys(Curr_Return_Month)
    range = driver.find_element(By.XPATH, "//*[@id='ctl00_ddl_tax_payer_range']").send_keys(Mid_range)
    year = driver.find_element(By.ID, "ctl00_ddl_year").send_keys(Curr_Year)
    time.sleep(5)
    driver.find_element(By.XPATH, "//input[@id='ctl00_btn_go']").click()
    time.sleep(5)
    page_title = driver.title
    print(page_title)
    sheetname['b130'] = "Month::" + "" + Curr_Return_Month
    sheetname['b131'] = " Year::" + "" + Curr_Year
    sheetname['b132'] = "range::" + "" + Mid_range
    try:
        element = driver.find_element(By.XPATH, "//*[@id='ctl00_ContentPlaceHolder1_grd_main']")
        WebDriverWait(driver, 20).until(EC.visibility_of((element)))
        grd_main = driver.find_element(By.XPATH, "//*[@id='ctl00_ContentPlaceHolder1_grd_main']")
        grid_rows = grd_main.find_elements_by_tag_name("tr");
        row_count = len(grid_rows)
        if(row_count >= 2):
            Total_Count = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_lblcount").text
            print("::" +Total_Count)
            sheetname['b129'] = "NonFilersTotalGSTMiddle :: Loading fine !"
            sheetname['a129'] = "NonFilersTotalGSTMiddle "
            print("NonFilers Risk based - Page is loading fine !")
        elif (row_count == 1 ):
            error_sheet['b109'] = "NonFilersTotalGSTMiddle"
            error_sheet['c109'] = "No data found"
            error_sheet['b110'] = "Month::" + "" + Curr_Return_Month
            error_sheet['b111'] = " Year::" + "" + Curr_Year
            error_sheet['b112'] = "range::" + "" + Mid_range

    except Exception as e :
        print("The Exception found:", str(e))
        destination = "Media//NonFilersTotalGSTMiddle.png"
        driver.save_screenshot(destination)
        print("Error Screenshot saved as ::", destination)
        error_sheet['b109'] = "NonFilersTotalGSTMiddle"
        error_sheet['c109'] = "Error page occurs"
        error_sheet['b110'] = "Month::" + "" + Curr_Return_Month
        error_sheet['b111'] = " Year::" + "" + Curr_Year
        error_sheet['b112'] = "range::" + "" + Mid_range

    # Others .............
    driver.get(url + "/Reports/rptR3BTopNonFilers_Prev.aspx")
    time.sleep(2)
    driver.find_element(By.XPATH, "//input[@id= 'ctl00_rdb_taxpayer_type_2']").click()
    month = driver.find_element(By.XPATH, "//*[@id='ctl00_ddl_Month']").send_keys(Curr_Return_Month)
    range = driver.find_element(By.XPATH, "//*[@id='ctl00_ddl_tax_payer_range']").send_keys(Other_range)
    year = driver.find_element(By.ID, "ctl00_ddl_year").send_keys(Curr_Year)
    driver.find_element(By.XPATH, "//input[@id='ctl00_btn_go']").click()
    time.sleep(5)
    page_title = driver.title
    print(page_title)
    sheetname['b134'] = "Month::" + "" + Curr_Return_Month
    sheetname['b135'] = " Year::" + "" + Curr_Year
    sheetname['b136'] = "range::" + "" + Other_range
    try:
        element = driver.find_element(By.XPATH,"//*[@id='ctl00_ContentPlaceHolder1_grd_main']")
        WebDriverWait(driver, 20).until(EC.visibility_of((element)))
        grd_main = driver.find_element(By.XPATH,"//*[@id='ctl00_ContentPlaceHolder1_grd_main']")
        grid_rows = grd_main.find_elements_by_tag_name("tr");
        row_count = len(grid_rows)
        if (row_count >= 2):
            Total_Count = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_lblcount").text
            print("::" +Total_Count)
            print("NonFilers  Total GST other- Page is loading fine !")
            sheetname['b133'] ="NonFilersTotalGST :: Loading fine !"
            sheetname['a133'] ="NonFilersTotalGST"
        elif (row_count == 1):
            error_sheet['b110'] = "NonFilersTotalGST_Others"
            error_sheet['c110'] = "No data found"
            error_sheet['b111'] = "Month::" + "" + Curr_Return_Month
            error_sheet['b112'] = " Year::" + "" + Curr_Year
            error_sheet['b113'] = "range::" + "" + Other_range

    except Exception as e :
        print("The Exception found:", str(e))
        destination = "Media//NonFilersTotalGST_other.png"
        driver.save_screenshot(destination)
        print("Error Screenshot saved as ::", destination)
        error_sheet['b110'] = "NonFilersTotalGST_Others"
        error_sheet['c110'] = "Error page occurs"
        error_sheet['b111'] = "Month::" + "" + Curr_Return_Month
        error_sheet['b112'] = " Year::" + "" + Curr_Year
        error_sheet['b113'] = "range::" + "" + Other_range
    finally:
         print("Done: Non Filers Total GST...... ")
    # ..GSTR9..#.......
    time.sleep(3)
    driver.get(url + "/Reports/rptR3BGstr9.aspx")
    time.sleep(2)
    select = Select(driver.find_element(By.ID, 'ctl00_ddlGstr9'))
    select.select_by_visible_text("2018-19")
    range = driver.find_element(By.XPATH, "//*[@id='ctl00_ddl_tax_payer_range']").send_keys(Grp_range)
    driver.find_element(By.XPATH, "//input[@id='ctl00_btn_go']").click()
    time.sleep(5)
    page_title = driver.title
    print(page_title)
    sheetname['a137'] = "Nonfilers GSTR9"
    sheetname['b138'] = "2018-2019"
    sheetname['b139'] = "range::" + "" + Grp_range

    try:
        grid_xpath = "//*[@id='ctl00_ContentPlaceHolder1_grd_main']"
        element = driver.find_element(By.XPATH, grid_xpath)
        WebDriverWait(driver, 20).until(EC.visibility_of((element)))
        Total_Count = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_lblcount").text
        print("::" +Total_Count)
        grd_main = driver.find_element(By.XPATH, grid_xpath)
        grid_rows = grd_main.find_elements_by_tag_name("tr");
        row_count = len(grid_rows)
        if (row_count >= 2):
            sheetname['b137'] = "GSTR9Top :: Loading fine !"
            print("GSTR9 top Page is loading fine !")
        elif (row_count == 1 ):
            error_sheet['b114'] = "NonFilersTotalGST_Others"
            error_sheet['c114'] = "No data found"
            error_sheet['b115'] = "2018-2019"
            error_sheet['b116'] = "range::" + "" + Grp_range
    except Exception as e:
        print("The Exception found:", str(e))
        destination = "Media//GSTR9Top.png"
        driver.save_screenshot(destination)
        print("Error Screenshot saved as ::", destination)
        error_sheet['b114'] = "NonFilersTotalGST_Others"
        error_sheet['c114'] = "Error page occurs"
        error_sheet['b115'] = "2018-2019"
        error_sheet['b116'] = "range::" + "" + Grp_range

    # middle....
    time.sleep(5)
    driver.get(url + "/Reports/rptR3BGstr9.aspx")
    time.sleep(5)
    driver.find_element(By.XPATH, "//input[@id='ctl00_rdb_taxpayer_type_1']").click()
    time.sleep(5)
    select = Select(driver.find_element(By.ID, 'ctl00_ddlGstr9'))
    select.select_by_visible_text("2018-19")
    range = driver.find_element(By.XPATH, "//*[@id='ctl00_ddl_tax_payer_range']").send_keys(Grp_range)
    driver.find_element(By.XPATH, "//input[@id='ctl00_btn_go']").click()
    time.sleep(5)
    page_title = driver.title
    print(page_title)
    sheetname['a140'] = "GSTR9  middle"
    sheetname['b141'] ="Year ::"  + "" + "2018-2019"
    sheetname['b142'] = "range::" + "" + Grp_range
    try:
        grid_xpath = "//*[@id='ctl00_ContentPlaceHolder1_grd_main']"
        element = driver.find_element(By.XPATH, grid_xpath)
        WebDriverWait(driver, 20).until(EC.visibility_of((element)))
        Total_Count = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_lblcount").text
        print("::" +Total_Count)
        grd_main = driver.find_element(By.XPATH, grid_xpath)
        grid_rows = grd_main.find_elements_by_tag_name("tr");
        row_count = len(grid_rows)
        if (row_count >= 2):
            sheetname['b140'] = "GSTR9Middle :: Loading fine !"

            print("GSTR9 middle Page is loading fine !")
        elif (row_count == 1 ):
            error_sheet['b117'] = "GSTR9Middle"
            error_sheet['c117'] = "No data found"
            error_sheet['b118'] = "Year ::" + "" + "2018-2019"
            error_sheet['b119'] = "range::" + "" + Grp_range
    except Exception as e:
        print("The Exception found:", str(e))
        destination = "Media//GSTR9Middle.png"
        driver.save_screenshot(destination)
        print("Error Screenshot saved as ::", destination)
        error_sheet['b117'] = "GSTR9Middle"
        error_sheet['c117'] = "Error page occurs"
        error_sheet['b118'] = "Year ::" + "" + "2018-2019"
        error_sheet['b119'] = "range::" + "" + Grp_range
        print("Done : rptR3BGstr9_GSTR9......")
finally:
    srcfile.save('GstPrime_AnalyticReports_inputvalues_result - '+uid+'-'+ state +'.xls')
    print("End of Test")
    driver.close()

# .............................................................