import unittest
import time
import xlrd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys


class Assessment(unittest.TestCase):
    @classmethod
    def setUp(inst):
        # create a new Browser session
        inst.driver = webdriver.Chrome("/home/vinod/chromedriver")
        inst.driver.implicitly_wait(30)
        inst.driver.maximize_window()
        # navigate to the application home page
        inst.driver.get("http://amsin.hirepro.in")
        print ('\nEntered URL in browser')
        return inst.driver


    def test_Downloadscoresheet(self):
        # get the search textbox
        print ('User reached on project selection screen')
        self.driver.find_element_by_id("crpo").click()
        print ('Clicked on CRPO project')
        self.driver.find_element_by_name("alias").send_keys("accenturetest")
        print ('Entered Tenant alias "accenturetest"')
        self.driver.find_element_by_xpath("//*[@class='btn btn-default']").click()
        print ('Clicked on "Next" button to move on next screen')
        time.sleep(3)
        self.driver.find_element_by_xpath("//div[5]/div/div/div[2]/div/div[2]/button").click()
        print ('Clicked to "Vendors/TPO/Placecom" button')
        time.sleep(3)
        self.LoginName_field = self.driver.find_element_by_xpath("//div[2]/section/div[1]/div[2]/form/div[1]/input")
        self.Password_field = self.driver.find_element_by_xpath("//div[2]/section/div[1]/div[2]/form/div[2]/input")
        # enter search keyword and submit
        self.LoginName_field.send_keys("vinodkumar")
        print ('Entered Login name')
        self.Password_field.send_keys("Admin@123")
        print ('Entered Password')
        self.driver.find_element_by_xpath("//div[2]/section/div[1]/div[2]/form/div[4]/a").click()
        time.sleep(3)
        print('Clicked to "Login" button')

        Click_assessment_Module = self.driver.find_element_by_xpath("//*[@ui-sref='crpo.assessment']")
        Click_assessment_Module.click()
        print ('Clicked to assessment module')

### here change the test roll number ###

        selecttest = self.driver.find_element_by_xpath("//div[2]/section/div/div/div[2]/div/div/div[5]/div/div[1]/div[2]/div[1]/span[1]/input")
        selecttest.click()

        actn = self.driver.find_element_by_xpath("//div[2]/section/div/div/div[1]/div[1]/div[2]/form/div/button/spam")
        actn.click()

        upload_click = self.driver.find_element_by_xpath("//div[2]/section/div/div/div[1]/div[1]/div[2]/form/div/ul/li[9]/a")
        upload_click.click()
        time.sleep(2)

### here give the excel path from loacl system to upload file ###

        choosefile = self.driver.find_element_by_xpath("//*[@type='file']")
        choosefile.send_keys("/home/vinod/Downloads/uploadscore/109015493_27_12_10_6_48_CandidateTemplate2016_12_27_10_06_48.xls")

        uploafile = self.driver.find_element_by_xpath("//*[@ng-click='data.uploadScoreSheet(file);$hide();']")
        uploafile.click()

        time.sleep(3)

    @classmethod
    def tearDownClass(inst):
        # close the browser window
        inst.driver.quit()
        print ('\nBrowser Closed')


if __name__ == '__main__':
    unittest.main()
