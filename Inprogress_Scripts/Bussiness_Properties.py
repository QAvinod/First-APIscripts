import time
import json
import requests
import xlwt
import datetime
import xlrd


class UploadCandidate:
    def __init__(self):

        # ------------------------
        # CRPO LOGIN APPLICATION
        # ------------------------
        self.header = {"content-type": "application/json"}
        self.TenantAlias = raw_input('TenantAlias:: ')
        self.LoginName = raw_input('LoginName:: ')
        self.Password = raw_input('Password:: ')

        self.login_request = {"LoginName": self.LoginName,
                              "Password": self.Password,
                              "TenantAlias": self.TenantAlias,
                              "UserName": self.LoginName}
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
        self.xl_eventId = []  # [] Initialising data from excel sheet to the variables
        self.xl_jobRoleId = []
        self.xl_mjrId = []
        self.xl_testId = []
        self.xl_Name = []
        self.xl_FirstName = []
        self.xl_MiddleName = []
        self.xl_LastName = []
        self.xl_Mobile1 = []
        self.xl_Email1 = []
        self.xl_Gender = []
        self.xl_DateOfBirth = []
        self.xl_USN = []
        self.xl_FinalPercentage = []
        self.xl_FinalCollegeId = []
        self.xl_FinalDegreeId = []
        self.xl_FinalDegreeTypeId = []

        self.xl_SourceId = []
        # self.xl_CampusId = []
        # self.xl_SourceType = []

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

        # -------------------------------------
        # Excel sheet write for Output results
        # -------------------------------------
        now = datetime.datetime.now()
        self.__current_DateTime = now.strftime("%d-%m-%Y-%H-%M-%S")
        self.wb_Result = xlwt.Workbook()
        self.ws = self.wb_Result.add_sheet('Candidates')
        self.rowsize = 1
        self.size = self.rowsize
        self.col = 0

        index = 0
        excelheaders = ['Comparison', 'Candidate_Created', 'Candidate Id', 'Event Id', 'Event Name', 'Job Id',
                        'Job Name', 'Applicant Id', 'Test Id', 'Test Name', 'Original CId', 'Message', 'Name',
                        'FirstName', 'MiddleName', 'LastName', 'Mobile1', 'Email1', 'Gender', 'DateOfBirth', 'USN',
                        'Final%', 'FinalDegree', 'FinalCollege', 'FinalDegreeType', 'SourceId']
        for headers in excelheaders:
            if headers in ['Comparison', 'Candidate Id', 'Original CId', 'Event Id', 'Event Name', 'Job Id',
                           'Job Name', 'Applicant Id', 'Candidate_Created', 'Message', 'Test Id', 'Test Name']:
                self.ws.write(0, index, headers, self.__style2)
            else:
                self.ws.write(0, index, headers, self.__style0)
            index += 1

        # -----------------------------------------------------------------------------------------------
        # Dictionary for CandidateGetbyIdDetails, CandidateEducationalDetails, CandidateExperienceDetails
        # -----------------------------------------------------------------------------------------------
        self.personal_details_dict = {}
        self.candidate_personal_details = self.personal_details_dict
        self.source_details_dict = {}
        self.candidate_source_details = self.source_details_dict
        self.final_degree_dict = {}
        self.candidate_final_degree_dict = self.final_degree_dict
        self.app_dict = {}
        self.event_applicant_dict = self.app_dict
        self.test_dict = {}
        self.test_detail = self.test_dict

    def excel_data(self):
        # ----------------
        # Excel Data Read
        # ----------------
        workbook = xlrd.open_workbook('/home/vinod/Desktop/Input/UploadCandidateScenarios/Basic_Properties.xls')
        sheet1 = workbook.sheet_by_index(0)
        for i in range(1, sheet1.nrows):
            number = i  # Counting number of rows
            rows = sheet1.row_values(number)

            # ------------------------------
            # Event, Job, Mjr, Test details
            # ------------------------------
            self.xl_eventId.append(int(rows[0]))
            self.xl_jobRoleId.append(int(rows[1]))
            self.xl_mjrId.append(int(rows[2]))
            self.xl_testId.append(int(rows[3]))

            #  -----------------------------
            # Personal, Source, Educational
            # ------------------------------
            self.xl_Name.append(str(rows[4]))
            self.xl_FirstName.append(str(rows[5]))
            self.xl_MiddleName.append(str(rows[6]))
            self.xl_LastName.append(str(rows[7]))
            self.xl_Mobile1.append(int(rows[8]))
            self.xl_Email1.append(str(rows[9]))
            self.xl_Gender.append(int(rows[10]))
            self.xl_DateOfBirth.append(str(rows[11]))
            self.xl_USN.append(str(rows[12]))
            self.xl_FinalPercentage.append(float(rows[13]))
            self.xl_FinalDegreeId.append(int(rows[14]))
            self.xl_FinalCollegeId.append(int(rows[15]))
            self.xl_FinalDegreeTypeId.append(int(rows[16]))
            self.xl_SourceId.append(int(rows[17]))
            # self.xl_CampusId.append(int(rows([])))
            # self.xl_SourceType.append(int(rows([])))

    def bulkCreateTagCandidates(self, iteration):
        # -------------------------
        # Candidate create request
        # -------------------------
        self.create_candidate_request = {"createTagCandidates": [{
            "PersonalDetails": {
                "Name": self.xl_Name[iteration],
                "FirstName": self.xl_FirstName[iteration],
                "MiddleName": self.xl_MiddleName[iteration],
                "LastName": self.xl_LastName[iteration],
                "Mobile1": self.xl_Mobile1[iteration],
                "Email1": self.xl_Email1[iteration],
                "Gender": self.xl_Gender[iteration],
                "DateOfBirth": self.xl_DateOfBirth[iteration],
                "USN": self.xl_USN[iteration],

            },
            "EducationDetails": {
                "AddedItems": [{
                    "IsPercentage": True,
                    "Percentage": self.xl_FinalPercentage[iteration],
                    "IsFinal": True,
                    "DegreeId": self.xl_FinalDegreeId[iteration],
                    "CollegeId": self.xl_FinalCollegeId[iteration],
                    "DegreeTypeId": self.xl_FinalDegreeTypeId[iteration]
                }]},
            "SourceDetails": {
                "SourceId": self.xl_SourceId[iteration],
                #     "CampusId": self.xl_CampusId[iteration],
                #     "SourceType": self.xl_SourceType[iteration]
            },
            "applicantDetail": {
                "eventId": self.xl_eventId[iteration],
                "jobRoleId": self.xl_jobRoleId[iteration],
                "mjrId": self.xl_mjrId[iteration],
                "testId": [self.xl_testId[iteration]],
                "isCreateDuplicate": True
            }
        }],
            "Sync": "True"
        }
        create_candidate = requests.post("https://amsin.hirepro.in/py/crpo/candidate/api/v1/bulkCreateTagCandidates/",
                                         headers=self.get_token,
                                         data=json.dumps(self.create_candidate_request, default=str), verify=False)
        create_candidate_response_dict = json.loads(create_candidate.content)
        candidate_response_data = create_candidate_response_dict['data']
        print candidate_response_data
        # print createcandidate.headers

        # -----------------------------------------
        # API response from bulkCreateTagCandidate
        # -----------------------------------------
        for response in candidate_response_data:
            self.isCreated = response['isCreated']
            self.OrginalCID = response.get('originalCandidateId')
            self.message = response.get('duplicateCandidateMessage')
            self.CID = response.get('candidateId')

            if self.isCreated:  # Always Boolean is true
                print "Create Candidate :", self.isCreated
                print "candidate Id ::", self.CID
            else:
                print "Create Candidate ::", self.isCreated
                print "Message ::", self.message

    def CandidateGetbyIdDetails(self):
        get_candidate_details = requests.post("https://amsin.hirepro.in/py/rpo/get_candidate_by_id/{}/"
                                              .format(self.CID), headers=self.get_token)
        candidate_details = json.loads(get_candidate_details.content)
        candidate_dict = candidate_details['Candidate']
        self.personal_details_dict = candidate_dict['PersonalDetails']
        self.source_details_dict = candidate_dict['SourceDetails']

    def CandidateEducationalDetails(self, loop):
        get_educational_details = requests.post("https://amsin.hirepro.in/py/rpo/get_candidate_education_details/{}/"
                                                .format(self.CID), headers=self.get_token)
        educational_details = json.loads(get_educational_details.content)
        educational_dict = educational_details['EducationProfile']
        for edu in educational_dict:
            if edu['DegreeId'] == self.xl_FinalDegreeId[loop]:
                self.final_degree_dict = next(
                    (item for item in educational_dict if item['DegreeId'] == self.xl_FinalDegreeId[loop]), None)

    def Event_Applicants(self, loop):
        eventapplicant_request = {
            "RecruitEventId": self.xl_eventId[loop],
            "PagingCriteriaType": {
                "MaxResults": 1000,
                "PageNumber": 1
            }
        }
        eventapplicant_api = requests.post("https://amsin.hirepro.in/py/crpo/applicant/api/v1/getAllApplicants/",
                                           headers=self.get_token,
                                           data=json.dumps(eventapplicant_request, default=str), verify=False)
        applicant_dict = json.loads(eventapplicant_api.content)
        # print applicant_dict
        applicant_data = applicant_dict['data']
        # print applicant_data
        for appdata in applicant_data:
            if appdata['CandidateId'] == self.CID:
                self.app_dict = next((item for item in applicant_data if item['CandidateId'] == self.CID), None)
                test_details = self.app_dict['TestUserDetailType']
                print test_details
                for td in test_details:
                    if td['TestId'] == self.xl_testId[loop]:
                        self.test_dict = next(
                            (item for item in test_details if item['TestId'] == self.xl_testId[loop]), None)

    def output_excel(self, loop):

        # ------------------
        # Writing Input Data
        # ------------------
        self.ws.write(self.rowsize, self.col, 'Input', self.__style4)
        self.ws.write(self.rowsize, 12, self.xl_Name[loop], self.__style1)
        self.ws.write(self.rowsize, 13, self.xl_FirstName[loop], self.__style1)
        self.ws.write(self.rowsize, 14, self.xl_MiddleName[loop], self.__style1)
        self.ws.write(self.rowsize, 15, self.xl_LastName[loop], self.__style1)
        self.ws.write(self.rowsize, 16, self.xl_Mobile1[loop], self.__style1)
        self.ws.write(self.rowsize, 17, self.xl_Email1[loop], self.__style1)
        self.ws.write(self.rowsize, 18, self.xl_Gender[loop], self.__style1)
        self.ws.write(self.rowsize, 19, self.xl_DateOfBirth[loop], self.__style1)
        self.ws.write(self.rowsize, 20, self.xl_USN[loop], self.__style1)
        self.ws.write(self.rowsize, 21, self.xl_FinalPercentage[loop], self.__style1)
        self.ws.write(self.rowsize, 22, self.xl_FinalDegreeId[loop], self.__style1)
        self.ws.write(self.rowsize, 23, self.xl_FinalCollegeId[loop], self.__style1)
        self.ws.write(self.rowsize, 24, self.xl_FinalDegreeTypeId[loop], self.__style1)
        self.ws.write(self.rowsize, 25, self.xl_SourceId[loop], self.__style1)

        # -------------------
        # Writing Output data
        # -------------------
        self.rowsize += 1  # Row increment
        self.ws.write(self.rowsize, self.col, 'Output', self.__style5)
        self.ws.write(self.rowsize, 1, str(self.isCreated))
        self.ws.write(self.rowsize, 2, self.CID)
        self.ws.write(self.rowsize, 3, self.app_dict.get('EventId', None))
        self.ws.write(self.rowsize, 4, self.app_dict.get('EventName', None))
        self.ws.write(self.rowsize, 5, self.app_dict.get('JobId', None))
        self.ws.write(self.rowsize, 6, self.app_dict.get('JobName', None))
        self.ws.write(self.rowsize, 7, self.app_dict.get('ApplicantId', None))
        self.ws.write(self.rowsize, 8, self.test_dict.get('TestId', None))
        self.ws.write(self.rowsize, 9, self.test_dict.get('TestName', None))
        self.ws.write(self.rowsize, 10, self.OrginalCID, self.__style3)
        self.ws.write(self.rowsize, 11, self.message, self.__style3)

        # ------------------------------------------------------------------
        # Comparing API Data with Excel Data and Printing into Output Excel
        # ------------------------------------------------------------------
        if self.xl_Name[loop] == self.personal_details_dict.get('Name'):
            self.ws.write(self.rowsize, 12, self.personal_details_dict.get('Name'))
        else:
            self.ws.write(self.rowsize, 12, self.personal_details_dict.get('Name', 'Duplicate'), self.__style3)

        if self.xl_FirstName[loop] == self.personal_details_dict.get('FirstName'):
            self.ws.write(self.rowsize, 13, self.personal_details_dict.get('FirstName'))
        else:
            self.ws.write(self.rowsize, 13, self.personal_details_dict.get('FirstName', 'Duplicate'), self.__style3)

        if self.xl_MiddleName[loop] == self.personal_details_dict.get('MiddleName'):
            self.ws.write(self.rowsize, 14, self.personal_details_dict.get('MiddleName'))
        else:
            self.ws.write(self.rowsize, 14, self.personal_details_dict.get('MiddleName', 'Duplicate'), self.__style3)

        if self.xl_LastName[loop] == self.personal_details_dict.get('LastName'):
            self.ws.write(self.rowsize, 15, self.personal_details_dict.get('LastName'))
        else:
            self.ws.write(self.rowsize, 15, self.personal_details_dict.get('LastName', 'Duplicate'), self.__style3)

        if str(self.xl_Mobile1[loop]) == self.personal_details_dict.get('Mobile1'):
            self.ws.write(self.rowsize, 16, int(self.personal_details_dict.get('Mobile1')))
        else:
            self.ws.write(self.rowsize, 16, self.personal_details_dict.get('Mobile1', 'Duplicate'), self.__style3)

        if self.xl_Email1[loop] == self.personal_details_dict.get('Email1'):
            self.ws.write(self.rowsize, 17, self.personal_details_dict.get('Email1'))
        else:
            self.ws.write(self.rowsize, 17, self.personal_details_dict.get('Email1', 'Duplicate'), self.__style3)

        if self.xl_Gender[loop] == self.personal_details_dict.get('Gender'):
            self.ws.write(self.rowsize, 18, self.personal_details_dict.get('Gender'))
        else:
            self.ws.write(self.rowsize, 18, self.personal_details_dict.get('Gender', 'Duplicate'), self.__style3)

        if self.xl_DateOfBirth[loop] == self.personal_details_dict.get('DateOfBirth'):
            self.ws.write(self.rowsize, 19, self.personal_details_dict.get('DateOfBirth'))
        else:
            self.ws.write(self.rowsize, 19, self.personal_details_dict.get('DateOfBirth', 'Duplicate'), self.__style3)

        if self.xl_USN[loop] == self.personal_details_dict.get('USN'):
            self.ws.write(self.rowsize, 20, self.personal_details_dict.get('USN'))
        else:
            self.ws.write(self.rowsize, 20, self.personal_details_dict.get('USN', 'Duplicate'), self.__style3)

        if self.xl_FinalPercentage[loop] == self.final_degree_dict.get('Percentage'):
            self.ws.write(self.rowsize, 21, self.final_degree_dict.get('Percentage'))
        else:
            self.ws.write(self.rowsize, 21, self.final_degree_dict.get('Percentage', 'Duplicate'), self.__style3)

        if self.xl_FinalDegreeId[loop] == self.final_degree_dict.get('DegreeId'):
            self.ws.write(self.rowsize, 22, self.final_degree_dict.get('DegreeId'))
        else:
            self.ws.write(self.rowsize, 22, self.final_degree_dict.get('DegreeId', 'Duplicate'), self.__style3)

        if self.xl_FinalCollegeId[loop] == self.final_degree_dict.get('CollegeId'):
            self.ws.write(self.rowsize, 23, self.final_degree_dict.get('CollegeId'))
        else:
            self.ws.write(self.rowsize, 23, self.final_degree_dict.get('CollegeId', 'Duplicate'), self.__style3)

        if self.xl_FinalDegreeTypeId[loop] == self.final_degree_dict.get('DegreeTypeId'):
            self.ws.write(self.rowsize, 24, self.final_degree_dict.get('DegreeTypeId'))
        else:
            self.ws.write(self.rowsize, 24, self.final_degree_dict.get('DegreeTypeId', 'Duplicate'), self.__style3)

        if self.xl_SourceId[loop] == self.source_details_dict.get('SourceId'):
            self.ws.write(self.rowsize, 25, self.source_details_dict.get('SourceId'))
        else:
            self.ws.write(self.rowsize, 25, self.source_details_dict.get('SourceId', 'Duplicate'), self.__style3)

        self.rowsize += 1  # Row increment
        Obj.wb_Result.save('/home/vinod/Desktop/Output/API_UploadCandidates(' + self.__current_DateTime + ').xls')


Obj = UploadCandidate()
Obj.excel_data()
Total_count = len(Obj.xl_Name)
print "Number Of Rows ::", Total_count
if Obj.login == 'OK':
    for looping in range(0, Total_count):
        print "Iteration Count is ::", looping
        Obj.bulkCreateTagCandidates(looping)
        if Obj.isCreated:  # Always Boolean is true, if it is not mention
            Obj.CandidateGetbyIdDetails()
            Obj.CandidateEducationalDetails(looping)
            Obj.Event_Applicants(looping)
        Obj.output_excel(looping)
        Obj.personal_details_dict = {}
        Obj.source_details_dict = {}
        Obj.final_degree_dict = {}
        Obj.app_dict = {}
        Obj.test_dict = {}
