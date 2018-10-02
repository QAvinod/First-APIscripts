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
        # self.xl_eventId     = []
        # self.xl_jobRoleId   = []
        # self.xl_mjrId       = []
        # self.xl_testId      = []
        self.xl_Name = []  # [] Initialising data from excel sheet to the variables
        self.xl_FirstName = []
        self.xl_MiddleName = []
        self.xl_LastName = []
        self.xl_Mobile1 = []
        self.xl_PhoneOffice = []
        self.xl_Email1 = []
        self.xl_Email2 = []
        self.xl_Gender = []
        self.xl_MaritalStatus = []
        self.xl_DateOfBirth = []
        self.xl_USN = []
        self.xl_Address1 = []
        self.xl_Address2 = []
        self.xl_PanNo = []
        self.xl_PassportNo = []
        self.xl_CurrentLocationId = []
        self.xl_TotalExperienceInMonths = []
        self.xl_Country = []
        self.xl_HierarchyId = []
        self.xl_Nationality = []
        self.xl_Sensitivity = []
        self.xl_StatusId = []
        self.xl_FinalPercentage = []
        self.xl_FinalEndYear = []
        self.xl_FinalDegreeId = []
        self.xl_FinalCollegeId = []
        self.xl_FinalDegreeTypeId = []
        self.xl_10thDegreeId = []
        self.xl_10thPercentage = []
        self.xl_10thEndYear = []
        self.xl_12thDegreeId = []
        self.xl_12thPercentage = []
        self.xl_12thEndYear = []
        self.xl_SourceId = []
        self.xl_CampusId = []
        self.xl_SourceType = []
        self.xl_Experience = []
        self.xl_EmployerId = []
        self.xl_DesignationId = []
        self.xl_Expertise = []
        self.xl_NoticePeriod = []
        self.xl_Integer1 = []
        self.xl_Integer2 = []
        self.xl_Integer3 = []
        self.xl_Integer4 = []
        self.xl_Integer5 = []
        self.xl_Integer6 = []
        self.xl_Integer7 = []
        self.xl_Integer8 = []
        self.xl_Integer9 = []
        self.xl_Integer10 = []
        self.xl_Integer11 = []
        self.xl_Integer12 = []
        self.xl_Integer13 = []
        self.xl_Integer14 = []
        self.xl_Integer15 = []
        self.xl_Text1 = []
        self.xl_Text2 = []
        self.xl_Text3 = []
        self.xl_Text4 = []
        self.xl_Text5 = []
        self.xl_Text6 = []
        self.xl_Text7 = []
        self.xl_Text8 = []
        self.xl_Text9 = []
        self.xl_Text10 = []
        self.xl_Text11 = []
        self.xl_Text12 = []
        self.xl_Text13 = []
        self.xl_Text14 = []
        self.xl_Text15 = []
        self.xl_TextArea1 = []
        self.xl_TextArea2 = []
        self.xl_TextArea3 = []
        self.xl_TextArea4 = []

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
        excelheaders = ['Comparison', 'Candidate_Created', 'Candidate Id', 'Original CId', 'Message', 'Name',
                        'FirstName', 'MiddleName', 'LastName', 'Mobile1', 'PhoneOffice', 'Email1', 'Email2',
                        'Gender', 'MaritalStatus', 'DateOfBirth', 'USN', 'Address1', 'Address2', 'Final%',
                        'FinalEndYear', 'FinalDegree', 'FinalCollege', 'FinalDegreeType', '10th%', '10thEndYear',
                        '12th%', '12thEndYear', 'PanNo', 'PassportNo', 'CurrentLocation', 'TotalExperienceInMonths',
                        'Country', 'HierarchyId', 'Nationality', 'Sensitivity', 'StatusId', 'SourceId', 'CampusId',
                        'SourceType', 'Experience', 'EmployerId', 'DesignationId', 'Expertise', 'Notice Period',
                        'Integer1', 'Integer2', 'Integer3', 'Integer4', 'Integer5', 'Integer6', 'Integer7', 'Integer8',
                        'Integer9', 'Integer10', 'Integer11', 'Integer12', 'Integer13', 'Integer14', 'Integer15',
                        'Text1', 'Text2', 'Text3', 'Text4', 'Text5', 'Text6', 'Text7', 'Text8', 'Text9', 'Text10',
                        'Text11', 'Text12', 'Text13', 'Text14', 'Text15', 'TextArea1', 'TextArea2', 'TextArea3',
                        'TextArea4']
        for headers in excelheaders:
            if headers in ['Comparison', 'Candidate Id', 'Original CId', 'Candidate_Created', 'Message']:
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
        self.custom_details_dict = {}
        self.candidate_custom_details = self.custom_details_dict
        self.final_degree_dict = {}
        self.candidate_final_degree_dict = self.final_degree_dict
        self.tenth_dict = {}
        self.candidate_tenth_dict = self.tenth_dict
        self.twelfth_dict = {}
        self.candidate_twelfth_dict = self.twelfth_dict
        self.experience_dict = {}
        self.candidate_experience_dict = self.experience_dict

    def excel_data(self):
        # ----------------
        # Excel Data Read
        # ----------------
        workbook = xlrd.open_workbook('/home/vinod/SOFTWARE/InputFiles/UploadCandidateScenarios/Candidate_Upload.xls')
        sheet1 = workbook.sheet_by_index(0)
        for i in range(1, sheet1.nrows):
            number = i  # Counting number of rows
            rows = sheet1.row_values(number)

            # ------------------------------
            # Event, Job, Mjr, Test details
            # ------------------------------
            # self.xl_eventId.append(int(rows[0]))
            # self.xl_jobRoleId.append(int(rows[1]))
            # self.xl_mjrId.append(int(rows[2]))
            # self.xl_testId.append(int(rows[3]))

            #  ----------------------------------------------------------
            # Personal, Source, Educational, Experience, Custom details
            # ----------------------------------------------------------
            self.xl_Name.append(str(rows[0]))
            self.xl_FirstName.append(str(rows[1]))
            self.xl_MiddleName.append(str(rows[2]))
            self.xl_LastName.append(str(rows[3]))
            self.xl_Mobile1.append(int(rows[4]))
            self.xl_PhoneOffice.append(int(rows[5]))
            self.xl_Email1.append(str(rows[6]))
            self.xl_Email2.append(str(rows[7]))
            self.xl_Gender.append(int(rows[8]))
            self.xl_MaritalStatus.append(int(rows[9]))
            self.xl_DateOfBirth.append(str(rows[10]))
            self.xl_USN.append(str(rows[11]))
            self.xl_Address1.append(str(rows[12]))
            self.xl_Address2.append(str(rows[13]))
            self.xl_PanNo.append(str(rows[14]))
            self.xl_PassportNo.append(str(rows[15]))
            self.xl_CurrentLocationId.append(int(rows[16]))
            self.xl_TotalExperienceInMonths.append(float(rows[17]))
            self.xl_Country.append(int(rows[18]))
            self.xl_HierarchyId.append(int(rows[19]))
            self.xl_Nationality.append(int(rows[20]))
            self.xl_Sensitivity.append(int(rows[21]))
            self.xl_StatusId.append(int(rows[22]))
            self.xl_FinalPercentage.append(float(rows[23]))
            self.xl_FinalEndYear.append(int(rows[24]))
            self.xl_FinalDegreeId.append(int(rows[25]))
            self.xl_FinalCollegeId.append(int(rows[26]))
            self.xl_FinalDegreeTypeId.append(int(rows[27]))
            self.xl_10thDegreeId.append(int(rows[28]))
            self.xl_10thPercentage.append(float(rows[29]))
            self.xl_10thEndYear.append(int(rows[30]))
            self.xl_12thDegreeId.append(int(rows[31]))
            self.xl_12thPercentage.append(float(rows[32]))
            self.xl_12thEndYear.append(int(rows[33]))
            self.xl_SourceId.append(int(rows[34]))
            self.xl_CampusId.append(int(rows[35]))
            self.xl_SourceType.append(int(rows[36]))
            self.xl_Experience.append(int(rows[37]))
            self.xl_EmployerId.append(int(rows[38]))
            self.xl_DesignationId.append(int(rows[39]))
            self.xl_Expertise.append(int(rows[40]))
            self.xl_NoticePeriod.append(int(rows[41]))
            self.xl_Integer1.append(int(rows[42]))
            self.xl_Integer2.append(int(rows[43]))
            self.xl_Integer3.append(int(rows[44]))
            self.xl_Integer4.append(int(rows[45]))
            self.xl_Integer5.append(int(rows[46]))
            self.xl_Integer6.append(int(rows[47]))
            self.xl_Integer7.append(int(rows[48]))
            self.xl_Integer8.append(int(rows[49]))
            self.xl_Integer9.append(int(rows[50]))
            self.xl_Integer10.append(int(rows[51]))
            self.xl_Integer11.append(int(rows[52]))
            self.xl_Integer12.append(int(rows[53]))
            self.xl_Integer13.append(int(rows[54]))
            self.xl_Integer14.append(int(rows[55]))
            self.xl_Integer15.append(int(rows[56]))
            self.xl_Text1.append(str(rows[57]))
            self.xl_Text2.append(str(rows[58]))
            self.xl_Text3.append(str(rows[59]))
            self.xl_Text4.append(str(rows[60]))
            self.xl_Text5.append(str(rows[61]))
            self.xl_Text6.append(str(rows[62]))
            self.xl_Text7.append(str(rows[63]))
            self.xl_Text8.append(str(rows[64]))
            self.xl_Text9.append(str(rows[65]))
            self.xl_Text10.append(str(rows[66]))
            self.xl_Text11.append(str(rows[67]))
            self.xl_Text12.append(str(rows[68]))
            self.xl_Text13.append(str(rows[69]))
            self.xl_Text14.append(str(rows[70]))
            self.xl_Text15.append(str(rows[71]))
            self.xl_TextArea1.append(str(rows[72]))
            self.xl_TextArea2.append(str(rows[73]))
            self.xl_TextArea3.append(str(rows[74]))
            self.xl_TextArea4.append(str(rows[75]))

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
                "PhoneOffice": self.xl_PhoneOffice[iteration],
                "Email1": self.xl_Email1[iteration],
                "Email2": self.xl_Email2[iteration],
                "Gender": self.xl_Gender[iteration],
                "DateOfBirth": self.xl_DateOfBirth[iteration],
                "USN": self.xl_USN[iteration],
                "MaritalStatus": self.xl_MaritalStatus[iteration],
                "Address1": self.xl_Address1[iteration],
                "Address2": self.xl_Address2[iteration],
                "PanNo": self.xl_PanNo[iteration],
                "PassportNo": self.xl_PassportNo[iteration],
                "CurrentLocationId": self.xl_CurrentLocationId[iteration],
                "TotalExperienceInMonths": self.xl_TotalExperienceInMonths[iteration],
                "Country": self.xl_Country[iteration],
                "HierarchyId": self.xl_HierarchyId[iteration],
                "Nationality": self.xl_Nationality[iteration],
                "Sensitivity": self.xl_Sensitivity[iteration],
                "StatusId": self.xl_StatusId[iteration],
                "ExpertiseId1": self.xl_Expertise[iteration]
            },
            "EducationDetails": {
                "AddedItems": [{
                    "IsPercentage": True,
                    "Percentage": self.xl_FinalPercentage[iteration],
                    "EndYear": self.xl_FinalEndYear[iteration],
                    "IsFinal": True,
                    "DegreeId": self.xl_FinalDegreeId[iteration],
                    "CollegeId": self.xl_FinalCollegeId[iteration],
                    "DegreeTypeId": self.xl_FinalDegreeTypeId[iteration]
                }, {
                    "IsPercentage": True,
                    "DegreeId": self.xl_10thDegreeId[iteration],
                    "Percentage": self.xl_10thPercentage[iteration],
                    "EndYear": self.xl_10thEndYear[iteration],
                    "IsFinal": False
                }, {
                    "IsPercentage": False,
                    "DegreeId": self.xl_12thDegreeId[iteration],
                    "Percentage": self.xl_12thPercentage[iteration],
                    "EndYear": self.xl_12thEndYear[iteration],
                    "IsFinal": False
                }]},
            "ExperienceDetails": {
                "AddedItems": [{
                    "IsLatest": True,
                    "Experience": self.xl_Experience[iteration],
                    "EmployerId": self.xl_EmployerId[iteration],
                    "DesignationId": self.xl_DesignationId[iteration]
                }]
            },
            "CustomDetails": {
                "Integer1": self.xl_Integer1[iteration],
                "Integer2": self.xl_Integer2[iteration],
                "Integer3": self.xl_Integer3[iteration],
                "Integer4": self.xl_Integer4[iteration],
                "Integer5": self.xl_Integer5[iteration],
                "Integer6": self.xl_Integer6[iteration],
                "Integer7": self.xl_Integer7[iteration],
                "Integer8": self.xl_Integer8[iteration],
                "Integer9": self.xl_Integer9[iteration],
                "Integer10": self.xl_Integer10[iteration],
                "Integer11": self.xl_Integer11[iteration],
                "Integer12": self.xl_Integer12[iteration],
                "Integer13": self.xl_Integer13[iteration],
                "Integer14": self.xl_Integer14[iteration],
                "Integer15": self.xl_Integer15[iteration],
                "Text1": self.xl_Text1[iteration],
                "Text2": self.xl_Text2[iteration],
                "Text3": self.xl_Text3[iteration],
                "Text4": self.xl_Text4[iteration],
                "Text5": self.xl_Text5[iteration],
                "Text6": self.xl_Text6[iteration],
                "Text7": self.xl_Text7[iteration],
                "Text8": self.xl_Text8[iteration],
                "Text9": self.xl_Text9[iteration],
                "Text10": self.xl_Text10[iteration],
                "Text11": self.xl_Text11[iteration],
                "Text12": self.xl_Text12[iteration],
                "Text13": self.xl_Text13[iteration],
                "Text14": self.xl_Text14[iteration],
                "Text15": self.xl_Text15[iteration],
                "TextArea1": self.xl_TextArea1[iteration],
                "TextArea2": self.xl_TextArea2[iteration],
                "TextArea3": self.xl_TextArea3[iteration],
                "TextArea4": self.xl_TextArea4[iteration]
            },
            "PreferenceDetails": {
                "NoticePeriod": self.xl_NoticePeriod[iteration]
            },
            "SourceDetails": {
                "SourceId": self.xl_SourceId[iteration],
                "CampusId": self.xl_CampusId[iteration],
                "SourceType": self.xl_SourceType[iteration]
            },
            # "applicantDetail":{
            #     "eventId":self.xl_eventId,
            #     "jobRoleId":self.xl_jobRoleId,
            #     "mjrId":self.xl_mjrId,
            #     "testId":[self.xl_testId],
            #     "isCreateDuplicate":True
            # }
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
        self.custom_details_dict = candidate_dict['CustomDetails']

    def CandidateEducationalDetails(self, loop):
        get_educational_details = requests.post("https://amsin.hirepro.in/py/rpo/get_candidate_education_details/{}/"
                                                .format(self.CID), headers=self.get_token)
        educational_details = json.loads(get_educational_details.content)
        educational_dict = educational_details['EducationProfile']
        for edu in educational_dict:
            if edu['DegreeId'] == self.xl_FinalDegreeId[loop]:
                self.final_degree_dict = next(
                    (item for item in educational_dict if item['DegreeId'] == self.xl_FinalDegreeId[loop]), None)
            if edu['DegreeId'] == self.xl_10thDegreeId[loop]:
                self.tenth_dict = next(
                    (item for item in educational_dict if item['DegreeId'] == self.xl_10thDegreeId[loop]), None)
            if edu['DegreeId'] == self.xl_12thDegreeId[loop]:
                self.twelfth_dict = next(
                    (item for item in educational_dict if item['DegreeId'] == self.xl_12thDegreeId[loop]), None)

    def CandidateExperienceDetails(self):
        get_experience_details = requests.post("https://amsin.hirepro.in/py/rpo/get_candidate_experience_details/{}/"
                                               .format(self.CID), headers=self.get_token)
        experience_details = json.loads(get_experience_details.content)
        experience_dict = experience_details['WorkProfile']
        for exp in experience_dict:
            self.experience_dict = exp

    def output_excel(self, loop):

        # ------------------
        # Writing Input Data
        # ------------------
        self.ws.write(self.rowsize, self.col, 'Input', self.__style4)
        self.ws.write(self.rowsize, 5, self.xl_Name[loop], self.__style1)
        self.ws.write(self.rowsize, 6, self.xl_FirstName[loop], self.__style1)
        self.ws.write(self.rowsize, 7, self.xl_MiddleName[loop], self.__style1)
        self.ws.write(self.rowsize, 8, self.xl_LastName[loop], self.__style1)
        self.ws.write(self.rowsize, 9, self.xl_Mobile1[loop], self.__style1)
        self.ws.write(self.rowsize, 10, self.xl_PhoneOffice[loop], self.__style1)
        self.ws.write(self.rowsize, 11, self.xl_Email1[loop], self.__style1)
        self.ws.write(self.rowsize, 12, self.xl_Email2[loop], self.__style1)
        self.ws.write(self.rowsize, 13, self.xl_Gender[loop], self.__style1)
        self.ws.write(self.rowsize, 14, self.xl_MaritalStatus[loop], self.__style1)
        self.ws.write(self.rowsize, 15, self.xl_DateOfBirth[loop], self.__style1)
        self.ws.write(self.rowsize, 16, self.xl_USN[loop], self.__style1)
        self.ws.write(self.rowsize, 17, self.xl_Address1[loop], self.__style1)
        self.ws.write(self.rowsize, 18, self.xl_Address2[loop], self.__style1)
        self.ws.write(self.rowsize, 19, self.xl_FinalPercentage[loop], self.__style1)
        self.ws.write(self.rowsize, 20, self.xl_FinalEndYear[loop], self.__style1)
        self.ws.write(self.rowsize, 21, self.xl_FinalDegreeId[loop], self.__style1)
        self.ws.write(self.rowsize, 22, self.xl_FinalCollegeId[loop], self.__style1)
        self.ws.write(self.rowsize, 23, self.xl_FinalDegreeTypeId[loop], self.__style1)
        self.ws.write(self.rowsize, 24, self.xl_10thPercentage[loop], self.__style1)
        self.ws.write(self.rowsize, 25, self.xl_10thEndYear[loop], self.__style1)
        self.ws.write(self.rowsize, 26, self.xl_12thPercentage[loop], self.__style1)
        self.ws.write(self.rowsize, 27, self.xl_12thEndYear[loop], self.__style1)
        self.ws.write(self.rowsize, 28, self.xl_PanNo[loop], self.__style1)
        self.ws.write(self.rowsize, 29, self.xl_PassportNo[loop], self.__style1)
        self.ws.write(self.rowsize, 30, self.xl_CurrentLocationId[loop], self.__style1)
        self.ws.write(self.rowsize, 31, self.xl_TotalExperienceInMonths[loop], self.__style1)
        self.ws.write(self.rowsize, 32, self.xl_Country[loop], self.__style1)
        self.ws.write(self.rowsize, 33, self.xl_HierarchyId[loop], self.__style1)
        self.ws.write(self.rowsize, 34, self.xl_Nationality[loop], self.__style1)
        self.ws.write(self.rowsize, 35, self.xl_Sensitivity[loop], self.__style1)
        self.ws.write(self.rowsize, 36, self.xl_StatusId[loop], self.__style1)
        self.ws.write(self.rowsize, 37, self.xl_SourceId[loop], self.__style1)
        self.ws.write(self.rowsize, 38, self.xl_CampusId[loop], self.__style1)
        self.ws.write(self.rowsize, 39, self.xl_SourceType[loop], self.__style1)
        self.ws.write(self.rowsize, 40, self.xl_Experience[loop], self.__style1)
        self.ws.write(self.rowsize, 41, self.xl_EmployerId[loop], self.__style1)
        self.ws.write(self.rowsize, 42, self.xl_DesignationId[loop], self.__style1)
        self.ws.write(self.rowsize, 43, self.xl_Expertise[loop], self.__style1)
        self.ws.write(self.rowsize, 44, self.xl_NoticePeriod[loop], self.__style1)
        self.ws.write(self.rowsize, 45, self.xl_Integer1[loop], self.__style1)
        self.ws.write(self.rowsize, 46, self.xl_Integer2[loop], self.__style1)
        self.ws.write(self.rowsize, 47, self.xl_Integer3[loop], self.__style1)
        self.ws.write(self.rowsize, 48, self.xl_Integer4[loop], self.__style1)
        self.ws.write(self.rowsize, 49, self.xl_Integer5[loop], self.__style1)
        self.ws.write(self.rowsize, 50, self.xl_Integer6[loop], self.__style1)
        self.ws.write(self.rowsize, 51, self.xl_Integer7[loop], self.__style1)
        self.ws.write(self.rowsize, 52, self.xl_Integer8[loop], self.__style1)
        self.ws.write(self.rowsize, 53, self.xl_Integer9[loop], self.__style1)
        self.ws.write(self.rowsize, 54, self.xl_Integer10[loop], self.__style1)
        self.ws.write(self.rowsize, 55, self.xl_Integer11[loop], self.__style1)
        self.ws.write(self.rowsize, 56, self.xl_Integer12[loop], self.__style1)
        self.ws.write(self.rowsize, 57, self.xl_Integer13[loop], self.__style1)
        self.ws.write(self.rowsize, 58, self.xl_Integer14[loop], self.__style1)
        self.ws.write(self.rowsize, 59, self.xl_Integer15[loop], self.__style1)
        self.ws.write(self.rowsize, 60, self.xl_Text1[loop], self.__style1)
        self.ws.write(self.rowsize, 61, self.xl_Text2[loop], self.__style1)
        self.ws.write(self.rowsize, 62, self.xl_Text3[loop], self.__style1)
        self.ws.write(self.rowsize, 63, self.xl_Text4[loop], self.__style1)
        self.ws.write(self.rowsize, 64, self.xl_Text5[loop], self.__style1)
        self.ws.write(self.rowsize, 65, self.xl_Text6[loop], self.__style1)
        self.ws.write(self.rowsize, 66, self.xl_Text7[loop], self.__style1)
        self.ws.write(self.rowsize, 67, self.xl_Text8[loop], self.__style1)
        self.ws.write(self.rowsize, 68, self.xl_Text9[loop], self.__style1)
        self.ws.write(self.rowsize, 69, self.xl_Text10[loop], self.__style1)
        self.ws.write(self.rowsize, 70, self.xl_Text11[loop], self.__style1)
        self.ws.write(self.rowsize, 71, self.xl_Text12[loop], self.__style1)
        self.ws.write(self.rowsize, 72, self.xl_Text13[loop], self.__style1)
        self.ws.write(self.rowsize, 73, self.xl_Text14[loop], self.__style1)
        self.ws.write(self.rowsize, 74, self.xl_Text15[loop], self.__style1)
        self.ws.write(self.rowsize, 75, self.xl_TextArea1[loop], self.__style1)
        self.ws.write(self.rowsize, 76, self.xl_TextArea2[loop], self.__style1)
        self.ws.write(self.rowsize, 77, self.xl_TextArea3[loop], self.__style1)
        self.ws.write(self.rowsize, 78, self.xl_TextArea4[loop], self.__style1)

        # -------------------
        # Writing Output data
        # -------------------
        self.rowsize += 1  # Row increment
        self.ws.write(self.rowsize, self.col, 'Output', self.__style5)
        self.ws.write(self.rowsize, 1, str(self.isCreated))
        self.ws.write(self.rowsize, 2, self.CID)
        self.ws.write(self.rowsize, 3, self.OrginalCID, self.__style3)
        self.ws.write(self.rowsize, 4, self.message, self.__style3)

        # ------------------------------------------------------------------
        # Comparing API Data with Excel Data and Printing into Output Excel
        # ------------------------------------------------------------------
        if self.xl_Name[loop] == self.personal_details_dict.get('Name'):
            self.ws.write(self.rowsize, 5, self.personal_details_dict.get('Name'))
        else:
            self.ws.write(self.rowsize, 5, self.personal_details_dict.get('Name', 'Duplicate'), self.__style3)

        if self.xl_FirstName[loop] == self.personal_details_dict.get('FirstName'):
            self.ws.write(self.rowsize, 6, self.personal_details_dict.get('FirstName'))
        else:
            self.ws.write(self.rowsize, 6, self.personal_details_dict.get('FirstName', 'Duplicate'), self.__style3)

        if self.xl_MiddleName[loop] == self.personal_details_dict.get('MiddleName'):
            self.ws.write(self.rowsize, 7, self.personal_details_dict.get('MiddleName'))
        else:
            self.ws.write(self.rowsize, 7, self.personal_details_dict.get('MiddleName', 'Duplicate'), self.__style3)

        if self.xl_LastName[loop] == self.personal_details_dict.get('LastName'):
            self.ws.write(self.rowsize, 8, self.personal_details_dict.get('LastName'))
        else:
            self.ws.write(self.rowsize, 8, self.personal_details_dict.get('LastName', 'Duplicate'), self.__style3)

        if str(self.xl_Mobile1[loop]) == self.personal_details_dict.get('Mobile1'):
            self.ws.write(self.rowsize, 9, int(self.personal_details_dict.get('Mobile1')))
        else:
            self.ws.write(self.rowsize, 9, self.personal_details_dict.get('Mobile1', 'Duplicate'), self.__style3)

        if str(self.xl_PhoneOffice[loop]) == self.personal_details_dict.get('PhoneOffice'):
            self.ws.write(self.rowsize, 10, int(self.personal_details_dict.get('PhoneOffice')))
        else:
            self.ws.write(self.rowsize, 10, self.personal_details_dict.get('PhoneOffice', 'Duplicate'), self.__style3)

        if self.xl_Email1[loop] == self.personal_details_dict.get('Email1'):
            self.ws.write(self.rowsize, 11, self.personal_details_dict.get('Email1'))
        else:
            self.ws.write(self.rowsize, 11, self.personal_details_dict.get('Email1', 'Duplicate'), self.__style3)

        if self.xl_Email2[loop] == self.personal_details_dict.get('Email2'):
            self.ws.write(self.rowsize, 12, self.personal_details_dict.get('Email2'))
        else:
            self.ws.write(self.rowsize, 12, self.personal_details_dict.get('Email2', 'Duplicate'), self.__style3)

        if self.xl_Gender[loop] == self.personal_details_dict.get('Gender'):
            self.ws.write(self.rowsize, 13, self.personal_details_dict.get('Gender'))
        else:
            self.ws.write(self.rowsize, 13, self.personal_details_dict.get('Gender', 'Duplicate'), self.__style3)

        if self.xl_MaritalStatus[loop] == self.personal_details_dict.get('MaritalStatus'):
            self.ws.write(self.rowsize, 14, self.personal_details_dict.get('MaritalStatus'))
        else:
            self.ws.write(self.rowsize, 14, self.personal_details_dict.get('MaritalStatus', 'Duplicate'), self.__style3)

        if self.xl_DateOfBirth[loop] == self.personal_details_dict.get('DateOfBirth'):
            self.ws.write(self.rowsize, 15, self.personal_details_dict.get('DateOfBirth'))
        else:
            self.ws.write(self.rowsize, 15, self.personal_details_dict.get('DateOfBirth', 'Duplicate'), self.__style3)

        if self.xl_USN[loop] == self.personal_details_dict.get('USN'):
            self.ws.write(self.rowsize, 16, self.personal_details_dict.get('USN'))
        else:
            self.ws.write(self.rowsize, 16, self.personal_details_dict.get('USN', 'Duplicate'), self.__style3)

        if self.xl_Address1[loop] == self.personal_details_dict.get('Address1'):
            self.ws.write(self.rowsize, 17, self.personal_details_dict.get('Address1'))
        else:
            self.ws.write(self.rowsize, 17, self.personal_details_dict.get('Address2', 'Duplicate'), self.__style3)

        if self.xl_Address2[loop] == self.personal_details_dict.get('Address2'):
            self.ws.write(self.rowsize, 18, self.personal_details_dict.get('Address2'))
        else:
            self.ws.write(self.rowsize, 18, self.personal_details_dict.get('Address2', 'Duplicate'), self.__style3)

        if self.xl_FinalPercentage[loop] == self.final_degree_dict.get('Percentage'):
            self.ws.write(self.rowsize, 19, self.final_degree_dict.get('Percentage'))
        else:
            self.ws.write(self.rowsize, 19, self.final_degree_dict.get('Percentage', 'Duplicate'), self.__style3)

        if self.xl_FinalEndYear[loop] == self.final_degree_dict.get('EndYear'):
            self.ws.write(self.rowsize, 20, self.final_degree_dict.get('EndYear'))
        else:
            self.ws.write(self.rowsize, 20, self.final_degree_dict.get('EndYear', 'Duplicate'), self.__style3)

        if self.xl_FinalDegreeId[loop] == self.final_degree_dict.get('DegreeId'):
            self.ws.write(self.rowsize, 21, self.final_degree_dict.get('DegreeId'))
        else:
            self.ws.write(self.rowsize, 21, self.final_degree_dict.get('DegreeId', 'Duplicate'), self.__style3)

        if self.xl_FinalCollegeId[loop] == self.final_degree_dict.get('CollegeId'):
            self.ws.write(self.rowsize, 22, self.final_degree_dict.get('CollegeId'))
        else:
            self.ws.write(self.rowsize, 22, self.final_degree_dict.get('CollegeId', 'Duplicate'), self.__style3)

        if self.xl_FinalDegreeTypeId[loop] == self.final_degree_dict.get('DegreeTypeId'):
            self.ws.write(self.rowsize, 23, self.final_degree_dict.get('DegreeTypeId'))
        else:
            self.ws.write(self.rowsize, 23, self.final_degree_dict.get('DegreeTypeId', 'Duplicate'), self.__style3)

        if self.xl_10thPercentage[loop] == self.tenth_dict.get('Percentage'):
            self.ws.write(self.rowsize, 24, self.tenth_dict.get('Percentage'))
        else:
            self.ws.write(self.rowsize, 24, self.tenth_dict.get('Percentage', 'Duplicate'), self.__style3)

        if self.xl_10thEndYear[loop] == self.tenth_dict.get('EndYear'):
            self.ws.write(self.rowsize, 25, self.tenth_dict.get('EndYear'))
        else:
            self.ws.write(self.rowsize, 25, self.tenth_dict.get('EndYear', 'Duplicate'), self.__style3)

        if self.xl_12thPercentage[loop] == self.twelfth_dict.get('Percentage'):
            self.ws.write(self.rowsize, 26, self.twelfth_dict.get('Percentage'))
        else:
            self.ws.write(self.rowsize, 26, self.twelfth_dict.get('Percentage', 'Duplicate'), self.__style3)

        if self.xl_12thEndYear[loop] == self.twelfth_dict.get('EndYear'):
            self.ws.write(self.rowsize, 27, self.twelfth_dict.get('EndYear'))
        else:
            self.ws.write(self.rowsize, 27, self.twelfth_dict.get('EndYear', 'Duplicate'), self.__style3)

        if self.xl_PanNo[loop] == self.personal_details_dict.get('PanNo'):
            self.ws.write(self.rowsize, 28, self.personal_details_dict.get('PanNo'))
        else:
            self.ws.write(self.rowsize, 28, self.personal_details_dict.get('PanNo', 'Duplicate'), self.__style3)

        if self.xl_PassportNo[loop] == self.personal_details_dict.get('PassportNo'):
            self.ws.write(self.rowsize, 29, self.personal_details_dict.get('PassportNo'))
        else:
            self.ws.write(self.rowsize, 29, self.personal_details_dict.get('PassportNo', 'Duplicate'), self.__style3)

        if self.xl_CurrentLocationId[loop] == self.personal_details_dict.get('CurrentLocationId'):
            self.ws.write(self.rowsize, 30, self.personal_details_dict.get('CurrentLocationId'))
        else:
            self.ws.write(self.rowsize, 30, self.personal_details_dict.get('CurrentLocationId', 'Duplicate'),
                          self.__style3)

        if self.xl_TotalExperienceInMonths[loop] == self.personal_details_dict.get('TotalExperienceInYears'):
            self.ws.write(self.rowsize, 31,
                          '{}.{}'.format(self.personal_details_dict.get('TotalExperienceInYears'),
                                         self.personal_details_dict.get('TotalExperienceInMonths')))
        else:
            self.ws.write(self.rowsize, 31,
                          '{}.{}'.format(self.personal_details_dict.get('TotalExperienceInYears', 'Duplicate'),
                                         self.personal_details_dict.get('TotalExperienceInMonths', 'Duplicate'),
                                         '--Converting to Year(s) & Month(s)'), self.__style3)

        if self.xl_Country[loop] == self.personal_details_dict.get('Country'):
            self.ws.write(self.rowsize, 32, self.personal_details_dict.get('Country'))
        else:
            self.ws.write(self.rowsize, 32, self.personal_details_dict.get('Country', 'Duplicate'), self.__style3)

        if self.xl_HierarchyId[loop] == self.personal_details_dict.get('HierarchyId'):
            self.ws.write(self.rowsize, 33, self.personal_details_dict.get('HierarchyId'))
        else:
            self.ws.write(self.rowsize, 33, self.personal_details_dict.get('HierarchyId', 'Duplicate'), self.__style3)

        if self.xl_Nationality[loop] == self.personal_details_dict.get('Nationality'):
            self.ws.write(self.rowsize, 34, self.personal_details_dict.get('Nationality'))
        else:
            self.ws.write(self.rowsize, 34, self.personal_details_dict.get('Nationality', 'Duplicate'), self.__style3)

        if self.xl_Sensitivity[loop] == self.personal_details_dict.get('Sensitivity'):
            self.ws.write(self.rowsize, 35, self.personal_details_dict.get('Sensitivity'))
        else:
            self.ws.write(self.rowsize, 35, self.personal_details_dict.get('Sensitivity', 'Duplicate'), self.__style3)

        if self.xl_StatusId[loop] == self.personal_details_dict.get('StatusId'):
            self.ws.write(self.rowsize, 36, self.personal_details_dict.get('StatusId'))
        else:
            self.ws.write(self.rowsize, 36, self.personal_details_dict.get('StatusId', 'Duplicate'), self.__style3)

        if self.xl_SourceId[loop] == self.source_details_dict.get('SourceId'):
            self.ws.write(self.rowsize, 37, self.source_details_dict.get('SourceId'))
        else:
            self.ws.write(self.rowsize, 37, self.source_details_dict.get('SourceId', 'Duplicate'), self.__style3)

        if self.xl_CampusId[loop] == self.source_details_dict.get('CampusId'):
            self.ws.write(self.rowsize, 38, self.source_details_dict.get('CampusId'))
        else:
            self.ws.write(self.rowsize, 38, self.source_details_dict.get('CampusId', 'Duplicate'), self.__style3)

        if self.xl_SourceType[loop] == self.source_details_dict.get('SourceType'):
            self.ws.write(self.rowsize, 39, self.source_details_dict.get('SourceType'))
        else:
            self.ws.write(self.rowsize, 39, self.source_details_dict.get('SourceType', 'Duplicate'), self.__style3)

        if self.xl_Experience[loop] == self.experience_dict.get('Experience'):
            self.ws.write(self.rowsize, 40, self.experience_dict.get('Experience'))
        else:
            self.ws.write(self.rowsize, 40, self.experience_dict.get('Experience', 'Duplicate'), self.__style3)

        if self.xl_EmployerId[loop] == self.experience_dict.get('EmployerId'):
            self.ws.write(self.rowsize, 41, self.experience_dict.get('EmployerId'))
        else:
            self.ws.write(self.rowsize, 41, self.experience_dict.get('EmployerId', 'Duplicate'), self.__style3)

        if self.xl_DesignationId[loop] == self.experience_dict.get('DesignationId'):
            self.ws.write(self.rowsize, 42, self.experience_dict.get('DesignationId'))
        else:
            self.ws.write(self.rowsize, 42, self.experience_dict.get('DesignationId', 'Duplicate'), self.__style3)

        if self.xl_Expertise[loop] == self.personal_details_dict.get('ExpertiseId1'):
            self.ws.write(self.rowsize, 43, self.personal_details_dict.get('ExpertiseId1'))
        else:
            self.ws.write(self.rowsize, 43, self.personal_details_dict.get('ExpertiseId1', 'Duplicate'), self.__style3)

        if self.xl_NoticePeriod[loop] == self.personal_details_dict.get('NoticePeriod'):
            self.ws.write(self.rowsize, 44, self.personal_details_dict.get('NoticePeriod'))
        else:
            self.ws.write(self.rowsize, 44, self.personal_details_dict.get('NoticePeriod', 'Duplicate'), self.__style3)

        if self.xl_Integer1[loop] == self.custom_details_dict.get('Integer1'):
            self.ws.write(self.rowsize, 45, self.custom_details_dict.get('Integer1'))
        else:
            self.ws.write(self.rowsize, 45, self.custom_details_dict.get('Integer1', 'Duplicate'), self.__style3)

        if self.xl_Integer2[loop] == self.custom_details_dict.get('Integer2'):
            self.ws.write(self.rowsize, 46, self.custom_details_dict.get('Integer2'))
        else:
            self.ws.write(self.rowsize, 46, self.custom_details_dict.get('Integer2', 'Duplicate'), self.__style3)

        if self.xl_Integer3[loop] == self.custom_details_dict.get('Integer3'):
            self.ws.write(self.rowsize, 47, self.custom_details_dict.get('Integer3'))
        else:
            self.ws.write(self.rowsize, 47, self.custom_details_dict.get('Integer3', 'Duplicate'), self.__style3)

        if self.xl_Integer4[loop] == self.custom_details_dict.get('Integer4'):
            self.ws.write(self.rowsize, 48, self.custom_details_dict.get('Integer4'))
        else:
            self.ws.write(self.rowsize, 48, self.custom_details_dict.get('Integer4', 'Duplicate'), self.__style3)

        if self.xl_Integer5[loop] == self.custom_details_dict.get('Integer5'):
            self.ws.write(self.rowsize, 49, self.custom_details_dict.get('Integer5'))
        else:
            self.ws.write(self.rowsize, 49, self.custom_details_dict.get('Integer5', 'Duplicate'), self.__style3)

        if self.xl_Integer6[loop] == self.custom_details_dict.get('Integer6'):
            self.ws.write(self.rowsize, 50, self.custom_details_dict.get('Integer6'))
        else:
            self.ws.write(self.rowsize, 50, self.custom_details_dict.get('Integer6', 'Duplicate'), self.__style3)

        if self.xl_Integer7[loop] == self.custom_details_dict.get('Integer7'):
            self.ws.write(self.rowsize, 51, self.custom_details_dict.get('Integer7'))
        else:
            self.ws.write(self.rowsize, 51, self.custom_details_dict.get('Integer7', 'Duplicate'), self.__style3)

        if self.xl_Integer8[loop] == self.custom_details_dict.get('Integer8'):
            self.ws.write(self.rowsize, 52, self.custom_details_dict.get('Integer8'))
        else:
            self.ws.write(self.rowsize, 52, self.custom_details_dict.get('Integer8', 'Duplicate'), self.__style3)

        if self.xl_Integer9[loop] == self.custom_details_dict.get('Integer9'):
            self.ws.write(self.rowsize, 53, self.custom_details_dict.get('Integer9'))
        else:
            self.ws.write(self.rowsize, 53, self.custom_details_dict.get('Integer9', 'Duplicate'), self.__style3)

        if self.xl_Integer10[loop] == self.custom_details_dict.get('Integer10'):
            self.ws.write(self.rowsize, 54, self.custom_details_dict.get('Integer10'))
        else:
            self.ws.write(self.rowsize, 54, self.custom_details_dict.get('Integer10', 'Duplicate'), self.__style3)

        if self.xl_Integer11[loop] == self.custom_details_dict.get('Integer11'):
            self.ws.write(self.rowsize, 55, self.custom_details_dict.get('Integer11'))
        else:
            self.ws.write(self.rowsize, 55, self.custom_details_dict.get('Integer11', 'Duplicate'), self.__style3)

        if self.xl_Integer12[loop] == self.custom_details_dict.get('Integer12'):
            self.ws.write(self.rowsize, 56, self.custom_details_dict.get('Integer12'))
        else:
            self.ws.write(self.rowsize, 56, self.custom_details_dict.get('Integer12', 'Duplicate'), self.__style3)

        if self.xl_Integer13[loop] == self.custom_details_dict.get('Integer13'):
            self.ws.write(self.rowsize, 57, self.custom_details_dict.get('Integer13'))
        else:
            self.ws.write(self.rowsize, 57, self.custom_details_dict.get('Integer13', 'Duplicate'), self.__style3)

        if self.xl_Integer14[loop] == self.custom_details_dict.get('Integer14'):
            self.ws.write(self.rowsize, 58, self.custom_details_dict.get('Integer14'))
        else:
            self.ws.write(self.rowsize, 58, self.custom_details_dict.get('Integer14', 'Duplicate'), self.__style3)

        if self.xl_Integer15[loop] == self.custom_details_dict.get('Integer15'):
            self.ws.write(self.rowsize, 59, self.custom_details_dict.get('Integer15'))
        else:
            self.ws.write(self.rowsize, 59, self.custom_details_dict.get('Integer15', 'Duplicate'), self.__style3)

        if self.xl_Text1[loop] == self.custom_details_dict.get('Text1'):
            self.ws.write(self.rowsize, 60, self.custom_details_dict.get('Text1'))
        else:
            self.ws.write(self.rowsize, 60, self.custom_details_dict.get('Text1', 'Duplicate'), self.__style3)

        if self.xl_Text2[loop] == self.custom_details_dict.get('Text2'):
            self.ws.write(self.rowsize, 61, self.custom_details_dict.get('Text2'))
        else:
            self.ws.write(self.rowsize, 61, self.custom_details_dict.get('Text2', 'Duplicate'), self.__style3)

        if self.xl_Text3[loop] == self.custom_details_dict.get('Text3'):
            self.ws.write(self.rowsize, 62, self.custom_details_dict.get('Text3'))
        else:
            self.ws.write(self.rowsize, 62, self.custom_details_dict.get('Text3', 'Duplicate'), self.__style3)

        if self.xl_Text4[loop] == self.custom_details_dict.get('Text4'):
            self.ws.write(self.rowsize, 63, self.custom_details_dict.get('Text4'))
        else:
            self.ws.write(self.rowsize, 63, self.custom_details_dict.get('Text4', 'Duplicate'), self.__style3)

        if self.xl_Text5[loop] == self.custom_details_dict.get('Text5'):
            self.ws.write(self.rowsize, 64, self.custom_details_dict.get('Text5'))
        else:
            self.ws.write(self.rowsize, 64, self.custom_details_dict.get('Text5', 'Duplicate'), self.__style3)

        if self.xl_Text6[loop] == self.custom_details_dict.get('Text6'):
            self.ws.write(self.rowsize, 65, self.custom_details_dict.get('Text6'))
        else:
            self.ws.write(self.rowsize, 65, self.custom_details_dict.get('Text6', 'Duplicate'), self.__style3)

        if self.xl_Text7[loop] == self.custom_details_dict.get('Text7'):
            self.ws.write(self.rowsize, 66, self.custom_details_dict.get('Text7'))
        else:
            self.ws.write(self.rowsize, 66, self.custom_details_dict.get('Text7', 'Duplicate'), self.__style3)

        if self.xl_Text8[loop] == self.custom_details_dict.get('Text8'):
            self.ws.write(self.rowsize, 67, self.custom_details_dict.get('Text8'))
        else:
            self.ws.write(self.rowsize, 67, self.custom_details_dict.get('Text8', 'Duplicate'), self.__style3)

        if self.xl_Text9[loop] == self.custom_details_dict.get('Text9'):
            self.ws.write(self.rowsize, 68, self.custom_details_dict.get('Text9'))
        else:
            self.ws.write(self.rowsize, 68, self.custom_details_dict.get('Text9', 'Duplicate'), self.__style3)

        if self.xl_Text10[loop] == self.custom_details_dict.get('Text10'):
            self.ws.write(self.rowsize, 69, self.custom_details_dict.get('Text10'))
        else:
            self.ws.write(self.rowsize, 69, self.custom_details_dict.get('Text10', 'Duplicate'), self.__style3)

        if self.xl_Text11[loop] == self.custom_details_dict.get('Text11'):
            self.ws.write(self.rowsize, 70, self.custom_details_dict.get('Text11'))
        else:
            self.ws.write(self.rowsize, 70, self.custom_details_dict.get('Text11', 'Duplicate'), self.__style3)

        if self.xl_Text12[loop] == self.custom_details_dict.get('Text12'):
            self.ws.write(self.rowsize, 71, self.custom_details_dict.get('Text12'))
        else:
            self.ws.write(self.rowsize, 71, self.custom_details_dict.get('Text12', 'Duplicate'), self.__style3)

        if self.xl_Text13[loop] == self.custom_details_dict.get('Text13'):
            self.ws.write(self.rowsize, 72, self.custom_details_dict.get('Text13'))
        else:
            self.ws.write(self.rowsize, 72, self.custom_details_dict.get('Text13', 'Duplicate'), self.__style3)

        if self.xl_Text14[loop] == self.custom_details_dict.get('Text14'):
            self.ws.write(self.rowsize, 73, self.custom_details_dict.get('Text14'))
        else:
            self.ws.write(self.rowsize, 73, self.custom_details_dict.get('Text14', 'Duplicate'), self.__style3)

        if self.xl_Text15[loop] == self.custom_details_dict.get('Text15'):
            self.ws.write(self.rowsize, 74, self.custom_details_dict.get('Text15'))
        else:
            self.ws.write(self.rowsize, 74, self.custom_details_dict.get('Text15', 'Duplicate'), self.__style3)

        if self.xl_TextArea1[loop] == self.custom_details_dict.get('TextArea1'):
            self.ws.write(self.rowsize, 75, self.custom_details_dict.get('TextArea1'))
        else:
            self.ws.write(self.rowsize, 75, self.custom_details_dict.get('TextArea1', 'Duplicate'), self.__style3)

        if self.xl_TextArea2[loop] == self.custom_details_dict.get('TextArea2'):
            self.ws.write(self.rowsize, 76, self.custom_details_dict.get('TextArea2'))
        else:
            self.ws.write(self.rowsize, 76, self.custom_details_dict.get('TextArea2', 'Duplicate'), self.__style3)

        if self.xl_TextArea3[loop] == self.custom_details_dict.get('TextArea3'):
            self.ws.write(self.rowsize, 77, self.custom_details_dict.get('TextArea3'))
        else:
            self.ws.write(self.rowsize, 77, self.custom_details_dict.get('TextArea3', 'Duplicate'), self.__style3)

        if self.xl_TextArea4[loop] == self.custom_details_dict.get('TextArea4'):
            self.ws.write(self.rowsize, 78, self.custom_details_dict.get('TextArea4'))
        else:
            self.ws.write(self.rowsize, 78, self.custom_details_dict.get('TextArea4', 'Duplicate'), self.__style3)

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
            Obj.CandidateExperienceDetails()
        Obj.output_excel(looping)
        Obj.personal_details_dict = {}
        Obj.source_details_dict = {}
        Obj.custom_details_dict = {}
        Obj.final_degree_dict = {}
        Obj.tenth_dict = {}
        Obj.twelfth_dict = {}
        Obj.experience_dict = {}
