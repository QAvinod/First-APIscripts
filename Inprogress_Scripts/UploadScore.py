import json
import requests
import time
import xlrd
import xlwt
import datetime


class UploadScoresheet:
    def __init__(self):

        # ------------------------
        # CRPO LOGIN APPLICATION
        # ------------------------
        self.header = {"content-type": "application/json"}
        # self.TenantAlias = raw_input('TenantAlias:: ')
        # self.LoginName = raw_input('LoginName:: ')
        # self.Password = raw_input('Password:: ')

        self.login_request = {"LoginName": 'admin',
                              "Password": 'Mohi@12345',
                              "TenantAlias": 'accenturetest',
                              "UserName": 'admin'}
        # self.server = raw_input('Server:: ')

        login_api = requests.post("https://amsin.hirepro.in/py/common/user/login_user/",
                                  headers=self.header,
                                  data=json.dumps(self.login_request),
                                  verify=False)
        self.response = login_api.json()
        self.get_token = {"content-type": "application/json",
                          "X-AUTH-TOKEN": self.response.get("Token")}
        self.var = None
        time.sleep(1)
        resp_dict = json.loads(login_api.content)
        self.status = resp_dict['status']
        if self.status == 'OK':
            self.login = 'OK'
            print "Login successfully"
            print "Status is", self.status
            time.sleep(1)
        else:
            self.login = 'KO'
            print "Failed to login"
            print "Status is", self.status

        # --------------------------
        # Initialising Excel Data
        # --------------------------
        self.xl_candidateId = []  # [] Initialising data from excel sheet to the variables
        self.xl_testId = []
        self.xl_group1 = []
        self.xl_section1 = []
        self.xl_section1_1 = []
        self.xl_group2 = []
        self.xl_section2 = []
        self.xl_section2_1 = []
        self.xl_group3 = []
        self.xl_section3 = []
        self.xl_section3_1 = []

        # -------------------------------------------------------
        # Styles for Excel sheet Row, Column, Text - color, Font
        # -------------------------------------------------------
        self.__style0 = xlwt.easyxf('pattern: pattern solid, fore_colour pale_blue;'
                                    'font: name Arial, color black, bold on;')
        self.__style1 = xlwt.easyxf('pattern: pattern solid, fore_colour gray25;'
                                    'font: name Arial, color black, bold off;')
        self.__style2 = xlwt.easyxf('pattern: pattern solid, fore_colour green;'
                                    'font: name Arial, color yellow, bold on;')
        self.__style3 = xlwt.easyxf('font: name Arial, color red, bold on')
        self.__style4 = xlwt.easyxf('pattern: pattern solid, fore_colour indigo;'
                                    'font: name Arial, color gold, bold on;')
        self.__style5 = xlwt.easyxf('pattern: pattern solid, fore_colour sky_blue;'
                                    'font: name Arial, color brown, bold on;')
        self.__style6 = xlwt.easyxf('pattern: pattern solid, fore_colour yellow;'
                                    'font: name Arial, color black, bold on;')
        self.__style7 = xlwt.easyxf('font: name Arial, color green, bold on')
        self.__style8 = xlwt.easyxf('font: name Arial, color orange, bold on')
        self.__style9 = xlwt.easyxf('pattern: pattern solid, fore_colour light_green;'
                                    'font: name Arial, color brown, bold on;')

        # -------------------------------------
        # Excel sheet write for Output results
        # -------------------------------------
        now = datetime.datetime.now()
        self.__current_DateTime = now.strftime("%d-%m-%Y-%H-%M-%S")
        self.wb_Result = xlwt.Workbook()
        self.ws = self.wb_Result.add_sheet('ScoreSheet')
        self.rowsize = 1
        self.size = self.rowsize
        self.col = 0

        index = 0
        excelheaders = ['Comparision', 'Candidate Id', 'CandidateName', 'Email', 'Mode', 'TotalMarks', 'Group1',
                        'Section1', 'Section1.1', 'Group2', 'Section2', 'Section2.1', 'Group3', 'Section3', 'Section3.1']
        for headers in excelheaders:
            if headers in ['Comparision', 'Candidate Id', 'CandidateName', 'Email', 'Mobile']:
                self.ws.write(0, index, headers, self.__style2)
            elif headers in ['Mode', 'TotalMarks']:
                self.ws.write(0, index, headers, self.__style9)
            else:
                self.ws.write(0, index, headers, self.__style0)
            index += 1

        # -----------------------------------------------------------------------------------------------
        # Dictionary for CandidateGetbyIdDetails, CandidateEducationalDetails, CandidateExperienceDetails
        # -----------------------------------------------------------------------------------------------
        self.testuser_dict = {}
        self.testuser_details = self.testuser_dict
        self.test_details_dict = {}
        self.test_details = self.test_details_dict
        self.candidate_info_dict = {}
        self.can_info = self.candidate_info_dict

        self.group1_dict = {}
        self.group_one = self.group1_dict
        self.section1_dict = {}
        self.section_one = self.section1_dict = {}
        self.section1_1_dict = {}
        self.section_one_one = self.section1_1_dict = {}

        self.group2_dict = {}
        self.group_two = self.group2_dict
        self.section2_dict = {}
        self.section_two = self.section2_dict = {}
        self.section2_1_dict = {}
        self.section_two_one = self.section2_1_dict = {}

        self.group3_dict = {}
        self.group_three = self.group3_dict
        self.section3_dict = {}
        self.section_three = self.section3_dict = {}
        self.section3_1_dict = {}
        self.section_three_one = self.section3_1_dict = {}

    # def download_sheet(self):
    #
    #     # ---------------------
    #     # Download Score Sheet
    #     # ---------------------
    #     downloadsheetrequest = {
    #         "TestId": 834,
    #         "IsSection": False
    #     }
    #     downloadsheet_request = requests.post("https://amsin.hirepro.in/py/crpo/assessment/api"
    #                                           "/v1/downloadCandidatesScore/", headers=self.get_token,
    #                                           data=json.dumps(downloadsheetrequest, default=str), verify=False)
    #     download_api_dict = json.loads(downloadsheet_request.content)
    #     download_api_data = download_api_dict['data']
    #     download_link = download_api_data['fileUrl']
    #     print download_link

    # def filehandler(self):
    #     dd = {"filename": "/home/vinod/Downloads/CandidateTemplate_20082018224953.xlsx"}
    #     r = requests.post("https://amsin.hirepro.in/py/common/filehandler/api/v2/upload/.xlsx/15000/",
    #                       headers=self.get_token,
    #                       data=json.dumps(dd, default=str), verify=False)
    #     r_dict = json.loads(r.content)
    #     print r_dict

    # def convert_path(self):
    #
    #     # -----------------------------
    #     # Saving Local path to S3 Path
    #     # -----------------------------
    #     persistent_request = [{
    #         "relativePath": "accenturetest/assessmentScoreSheets",
    #         "origFileUrl": "https://s3-ap-southeast-1.amazonaws.com/ams-in-self-expiring-files/1-24h/accenturetest/"
    #                        "uploaded/ed420e24-7aa1-499b-87cd-b797bf18e684CandidateTemplate_09072018174308.xlsx",
    #         "isSync": True
    #     }]
    #     persistent_api = requests.post("https://amsin.hirepro.in/py/common/filehandler/api/v2/persistent-save/",
    #                                    headers=self.get_token,
    #                                    data=json.dumps(persistent_request, default=str), verify=False)
    #     persistent_api_dict = json.loads(persistent_api.content)
    #     print persistent_api_dict

    def excel_data(self):
        # ----------------
        # Excel Data Read
        # ----------------
        workbook = xlrd.open_workbook('/home/vinod/Desktop/Input/UploadScores.xls')
        sheet1 = workbook.sheet_by_index(0)
        for i in range(1, sheet1.nrows):
            number = i  # Counting number of rows
            rows = sheet1.row_values(number)

            self.xl_candidateId.append(int(rows[0]))
            self.xl_testId.append(int(rows[1]))
            self.xl_group1.append(int(rows[3]))
            self.xl_section1.append(int(rows[4]))
            self.xl_section1_1.append(int(rows[5]))
            self.xl_group2.append(int(rows[6]))
            self.xl_section2.append(int(rows[7]))
            self.xl_section2_1.append(int(rows[8]))
            self.xl_group3.append(int(rows[9]))
            self.xl_section3.append(int(rows[10]))
            self.xl_section3_1.append(int(rows[11]))

    # def upload_sheet(self, loop):

        # ------------------          ---------------------------------------
        # Upload Score Sheet ******** Every 30 Days Replace the S3 "FilePath"
        # ------------------          ---------------------------------------
        # uploadsheetrequest = {
        #     "TestId": self.xl_testId[loop],
        #     "FilePath": "https://s3-ap-southeast-1.amazonaws.com/test-all-hirepro-files/accenturetest/"
        #                 "assessmentScoreSheets/"
        #                 "75b6439b-cb4e-4279-bf3e-d9a1d512ebcbCandidateTemplate_09082018190325.xlsx",
        #     "Sync": "False"
        # }
        # uploadsheet_api = requests.post("https://amsin.hirepro.in/py/crpo/assessment/api/v1/uploadCandidatesScore/",
        #                                 headers=self.get_token,
        #                                 data=json.dumps(uploadsheetrequest, default=str), verify=False)
        # upload_api_dict = json.loads(uploadsheet_api.content)
        # print upload_api_dict

    def fetching_scores(self, loop):
        score_request = {
            "CandidateIds": [self.xl_candidateId[loop]]
        }
        fetchingscores_api = requests.post("https://amsin.hirepro.in/py/crpo/applicant/api/v1/getApplicantsInfo/",
                                           headers=self.get_token,
                                           data=json.dumps(score_request, default=str), verify=False)
        fetchingscores_dict = json.loads(fetchingscores_api.content)
        scoredata = fetchingscores_dict['data']
        for testuser in scoredata:
            if testuser['CandidateId'] == self.xl_candidateId[loop]:
                self.testuser_dict = next(
                    (item for item in scoredata if item['CandidateId'] == self.xl_candidateId[loop]), None)
                # print testuser_dict

                assessment_dict = self.testuser_dict['AssessmentDetails']
                print assessment_dict
                for permission_to_go in assessment_dict:
                    self.is_offline = permission_to_go['IsOffline']
                    if permission_to_go['TestStatus'] == "NotAttended":
                        self.candidate_info_dict = next(
                            (item for item in assessment_dict if item['Id'] == self.xl_testId[loop]), None)
                    else:

                        for test_user in assessment_dict:
                            if test_user['Id'] == self.xl_testId[loop]:
                                self.test_details_dict = next(
                                    (item for item in assessment_dict if item['Id'] == self.xl_testId[loop]), None)
                                print self.test_details_dict
                                for group_details in self.test_details_dict['GroupWiseInfo']:

                                    # ------------
                                    # Group - 1
                                    # ------------
                                    if group_details['GroupId'] == self.xl_group1[loop]:
                                        self.group1_dict = next(
                                            (item for item in
                                             self.test_details_dict['GroupWiseInfo']
                                             if item['GroupId'] == self.xl_group1[loop]), None)
                                        self.section1_dict = next(
                                            (item for item in
                                             self.group1_dict['SectionInfoTypes']
                                             if item['SectionId'] == self.xl_section1[loop]), None)
                                        self.section1_1_dict = next(
                                            (item for item in
                                             self.group1_dict['SectionInfoTypes']
                                             if item['SectionId'] == self.xl_section1_1[loop]), None)
                                        # print self.section1_dict
                                        # print self.section1_1_dict
                                        # print self.group1_dict

                                    # ------------
                                    # Group - 2
                                    # ------------
                                    if group_details['GroupId'] == self.xl_group2[loop]:
                                        self.group2_dict = next(
                                            (item for item in
                                             self.test_details_dict['GroupWiseInfo']
                                             if item['GroupId'] == self.xl_group2[loop]), None)
                                        self.section2_dict = next(
                                            (item for item in
                                             self.group2_dict['SectionInfoTypes']
                                             if item['SectionId'] == self.xl_section2[loop]), None)
                                        self.section2_1_dict = next(
                                            (item for item in
                                             self.group2_dict['SectionInfoTypes']
                                             if item['SectionId'] == self.xl_section2_1[loop]), None)
                                        # print self.section2_dict
                                        # print self.section2_1_dict
                                        # print self.group2_dict

                                    # ------------
                                    # Group - 3
                                    # ------------
                                    if group_details['GroupId'] == self.xl_group3[loop]:
                                        self.group3_dict = next(
                                            (item for item in
                                             self.test_details_dict['GroupWiseInfo']
                                             if item['GroupId'] == self.xl_group3[loop]), None)
                                        self.section3_dict = next(
                                            (item for item in
                                             self.group3_dict['SectionInfoTypes']
                                             if item['SectionId'] == self.xl_section3[loop]), None)
                                        self.section3_1_dict = next(
                                            (item for item in
                                             self.group3_dict['SectionInfoTypes']
                                             if item['SectionId'] == self.xl_section3_1[loop]), None)
                                        # print self.section3_dict
                                        # print self.section3_1_dict
                                        # print self.group3_dict

    def output_excel(self):
        # ------------------
        # Writing Input Data
        # ------------------
        self.ws.write(self.rowsize, self.col, 'Input', self.__style4)

        # -------------------
        # Writing Output data
        # -------------------
        self.rowsize += 1  # Row increment
        self.ws.write(self.rowsize, self.col, 'Output', self.__style5)
        if self.testuser_dict and self.testuser_dict.get('CandidateId'):
            self.ws.write(self.rowsize, 1, self.testuser_dict.get('CandidateId', None))

        if self.testuser_dict and self.testuser_dict.get('CandidateName'):
            self.ws.write(self.rowsize, 2, self.testuser_dict.get('CandidateName', None))

        if self.testuser_dict and self.testuser_dict.get('Email'):
            self.ws.write(self.rowsize, 3, self.testuser_dict.get('Email', None))

        if self.is_offline:
            self.ws.write(self.rowsize, 4, "Offline", self.__style8)
        else:
            self.ws.write(self.rowsize, 4, "Online", self.__style7)

        if self.test_details_dict and self.test_details_dict.get('CandidateMarks'):
            self.ws.write(self.rowsize, 5, self.test_details_dict.get('CandidateMarks', None))
        else:
            self.ws.write(self.rowsize, 5, self.candidate_info_dict.get('TestStatus'), self.__style3)

        if self.group1_dict and self.group1_dict.get('CandidateScoreTotal'):
            self.ws.write(self.rowsize, 6, self.group1_dict.get('CandidateScoreTotal', None))

        if self.section1_dict and self.section1_dict.get('CandidateScoreTotal'):
            self.ws.write(self.rowsize, 7, self.section1_dict.get('CandidateScoreTotal', None))

        if self.section1_1_dict and self.section1_1_dict.get('CandidateScoreTotal'):
            self.ws.write(self.rowsize, 8, self.section1_1_dict.get('CandidateScoreTotal', None))

        if self.group2_dict and self.group2_dict.get('CandidateScoreTotal'):
            self.ws.write(self.rowsize, 9, self.group2_dict.get('CandidateScoreTotal', None))

        if self.section2_dict and self.section2_dict.get('CandidateScoreTotal'):
            self.ws.write(self.rowsize, 10, self.section2_dict.get('CandidateScoreTotal', None))

        if self.section2_1_dict and self.section2_1_dict.get('CandidateScoreTotal'):
            self.ws.write(self.rowsize, 11, self.section2_1_dict.get('CandidateScoreTotal', None))

        if self.group3_dict and self.group3_dict.get('CandidateScoreTotal'):
            self.ws.write(self.rowsize, 12, self.group3_dict.get('CandidateScoreTotal', None))

        if self.section3_dict and self.section3_dict.get('CandidateScoreTotal'):
            self.ws.write(self.rowsize, 13, self.section3_dict.get('CandidateScoreTotal', None))

        if self.section3_1_dict and self.section3_1_dict.get('CandidateScoreTotal'):
            self.ws.write(self.rowsize, 14, self.section3_1_dict.get('CandidateScoreTotal', None))

        self.rowsize += 1  # Row increment
        Object.wb_Result.save('/home/vinod/Desktop/Output/API_Download_upload_Scores.xls')


Object = UploadScoresheet()
Object.excel_data()

Total_count = len(Object.xl_candidateId)
print "Number Of Rows ::", Total_count
if Object.login == 'OK':
    for looping in range(0, Total_count):
        print "Iteration Count is ::", looping
        # Object.upload_sheet(looping)
        Object.fetching_scores(looping)
        Object.output_excel()
        Object.candidate_info_dict = {}
        Object.testuser_dict = {}
        Object.test_details_dict = {}
        Object.group1_dict = {}
        Object.group2_dict = {}
        Object.group3_dict = {}
        Object.section1_dict = {}
        Object.section2_dict = {}
        Object.section3_dict = {}
        Object.section1_1_dict = {}
        Object.section2_1_dict = {}
        Object.section3_1_dict = {}
