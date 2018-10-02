import time
import json
import requests
import xlwt
import xlrd
import datetime


class InterviewFeedback:

    def __init__(self):
        # ---------------------
        # CRPO LOGIN APPLICATION
        # ---------------------
        self.header = {"content-type": "application/json"}
        # self.TenantAlias = raw_input('TenantAlias:: ')
        # self.LoginName = raw_input('LoginName:: ')
        # self.Password = raw_input('Password:: ')

        self.login_request = {"LoginName": "admin",
                              "Password": "Mohi@1234",
                              "TenantAlias": "accenturetest",
                              "UserName": "admin"}
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
        self.__style6 = xlwt.easyxf('font: name Arial, color light_orange, bold on')
        self.__style7 = xlwt.easyxf('font: name Arial, color orange, bold on')
        self.__style8 = xlwt.easyxf('font: name Arial, color green, bold on')

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
        excelheaders = []
        for headers in excelheaders:
            if headers in []:
                self.ws.write(0, index, headers, self.__style2)
            else:
                self.ws.write(0, index, headers, self.__style0)
            index += 1

        # --------------------------
        # Initialising Excel Data
        # --------------------------
        self.xl_Event_id = []  # [] Initialising data from excel sheet to the variables
        self.xl_Applicant_id = []
        self.xl_Job_id = []
        self.xl_type = []
        self.xl_Datetime = []
        self.xl_stage_id = []
        self.xl_interviewers_id = []
        self.xl_Comment = []
        self.xl_location = []

        self.xl_Skill_id_01 = []
        self.xl_Skill_score_01 = []
        self.xl_Skill_id_02 = []
        self.xl_Skill_score_02 = []
        self.xl_Skill_id_03 = []
        self.xl_Skill_score_03 = []
        self.xl_Skill_id_04 = []
        self.xl_Skill_score_04 = []
        self.xl_skill_comment = []
        self.xl_decision = []
        self.xl_duration = []
        self.xl_int_datetime = []

        # -----------------------------------------------------------------------------------------------
        # Dictionary for CandidateGetbyIdDetails, CandidateEducationalDetails, CandidateExperienceDetails
        # -----------------------------------------------------------------------------------------------
        self.user_dict = {}
        self.user_get_details = self.user_dict

    def excel_data(self):
        # ----------------
        # Excel Data Read
        # ----------------
        try:
            workbook = xlrd.open_workbook('/home/vinod/SOFTWARE/InputFiles/GiveFeedback/GiveFeedback.xls')
            sheet = workbook.sheet_by_index(0)
            for i in range(1, sheet.nrows):
                number = i
                rows = sheet.row_values(number)

                if rows[0] is not None and rows[0] != '':
                    self.xl_Event_id.append(int(rows[0]))
                else:
                    self.xl_Event_id.append(None)

                if rows[1] is not None and rows[1] != '':
                    self.xl_Applicant_id.append(int(rows[1]))
                else:
                    self.xl_Applicant_id.append(None)

                if rows[2] is not None and rows[2] != '':
                    self.xl_Job_id.append(int(rows[2]))
                else:
                    self.xl_Job_id.append(None)

                if rows[3] is not None and rows[3] != '':
                    self.xl_type.append(int(rows[3]))
                else:
                    self.xl_type.append(None)

                if rows[4] is not None and rows[4] != '':
                    self.xl_Datetime.append(str(rows[4]))
                else:
                    self.xl_Datetime.append(None)

                if rows[5] is not None and rows[5] != '':
                    self.xl_stage_id.append(int(rows[5]))
                else:
                    self.xl_stage_id.append(None)

                if rows[6] is not None and rows[6] != '':
                    int_ids = map(int, rows[6].split(',') if isinstance(rows[6], basestring) else [rows[6]])
                    self.xl_interviewers_id.append(int_ids)
                else:
                    self.xl_interviewers_id.append(None)

                if rows[7] is not None and rows[7] != '':
                    self.xl_Comment.append(str(rows[7]))
                else:
                    self.xl_Comment.append(None)

                if rows[8] is not None and rows[8] != '':
                    self.xl_location.append(int(rows[8]))
                else:
                    self.xl_location.append(None)

                if rows[9] is not None and rows[9] != '':
                    self.xl_Skill_id_01.append(int(rows[9]))
                else:
                    self.xl_Skill_id_01.append(None)

                if rows[10] is not None and rows[10] != '':
                    self.xl_Skill_score_01.append(int(rows[10]))
                else:
                    self.xl_Skill_score_01.append(None)

                if rows[11] is not None and rows[11] != '':
                    self.xl_Skill_id_02.append(int(rows[11]))
                else:
                    self.xl_Skill_id_02.append(None)

                if rows[12] is not None and rows[12] != '':
                    self.xl_Skill_score_02.append(int(rows[12]))
                else:
                    self.xl_Skill_score_02.append(None)

                if rows[13] is not None and rows[13] != '':
                    self.xl_Skill_id_03.append(int(rows[13]))
                else:
                    self.xl_Skill_id_03.append(None)

                if rows[14] is not None and rows[14] != '':
                    self.xl_Skill_score_03.append(int(rows[14]))
                else:
                    self.xl_Skill_score_03.append(None)

                if rows[15] is not None and rows[15] != '':
                    self.xl_Skill_id_04.append(int(rows[15]))
                else:
                    self.xl_Skill_id_04.append(None)

                if rows[16] is not None and rows[16] != '':
                    self.xl_Skill_score_04.append(int(rows[16]))
                else:
                    self.xl_Skill_score_04.append(None)

                if rows[17] is not None and rows[17] != '':
                    self.xl_skill_comment.append(str(rows[17]))
                else:
                    self.xl_skill_comment.append(None)

                if rows[18] is not None and rows[18] != '':
                    self.xl_decision.append(int(rows[18]))
                else:
                    self.xl_decision.append(None)

                if rows[19] is not None and rows[19] != '':
                    self.xl_duration.append(int(rows[19]))
                else:
                    self.xl_duration.append(None)

                if rows[20] is not None and rows[20] != '':
                    self.xl_int_datetime.append(str(rows[20]))
                else:
                    self.xl_int_datetime.append(None)

        except IOError:
            print("File not found or path is incorrect")

    def schedule_interview(self, loop):
        schedule_request = [{
            "isConsultantRound": False,
            "interviewDate": self.xl_Datetime[loop],
            "interviewTime": "",
            "interviewType": self.xl_type[loop],
            "interviewerIds": self.xl_interviewers_id[loop],
            "jobId": self.xl_Job_id[loop],
            "stageId": self.xl_stage_id[loop],
            "locationId": self.xl_location[loop],  # API default send bangalore location
            "secondaryInterviewerIds": [],
            "recruiterComment": self.xl_Comment[loop],
            "recruitEventId": self.xl_Event_id[loop],
            "applicantIds": [self.xl_Applicant_id[loop]]
        }]
        scheduling_interviews = requests.post("https://amsin.hirepro.in/py/crpo/api/v1/interview/schedule/",
                                              headers=self.get_token,
                                              data=json.dumps(schedule_request, default=str), verify=False)
        schedule_response = json.loads(scheduling_interviews.content)
        print schedule_response
        data = schedule_response['data']

        if schedule_response['status'] == 'OK':
            success = data['success']
            failure = data['failure']
            print failure
            print success
        else:
            print ('Error occured while scheduling')

    def provide_feedback(self, loop):
        feedback_request = {
            "interviewRequestId": 65722,
            "interviewerFeedback": [{
                "skillsAssessed": [{
                    "skillId": self.xl_Skill_id_01[loop]
                }, {
                    "skillId": self.xl_Skill_id_02[loop],
                    "skillScore": self.xl_Skill_score_01[loop],
                    "skillComment": self.xl_skill_comment[loop]
                }, {
                    "skillId": self.xl_Skill_id_03[loop],
                    "skillScore": self.xl_Skill_score_03[loop],
                    "skillComment": self.xl_skill_comment[loop]
                }, {
                    "skillId": self.xl_Skill_id_04[loop],
                    "skillScore": self.xl_Skill_score_04[loop],
                    "skillComment": self.xl_skill_comment[loop]
                }],
                "interviwerIds": self.xl_interviewers_id[loop],
                "applicantId": self.xl_Applicant_id[loop],
                "interviewerDecision": self.xl_decision[loop],
                "interviewerComment": self.xl_skill_comment[loop],
                "interviewDuration": self.xl_duration[loop],
                "interviewedDate": self.xl_int_datetime[loop]
            }]
        }
        providing_feedback = requests.post("https://amsin.hirepro.in/py/crpo/api/v1/interview/schedule/",
                                           headers=self.get_token,
                                           data=json.dumps(feedback_request, default=str), verify=False)
        feedback_response = json.loads(providing_feedback.content)
        print feedback_response


Object = InterviewFeedback()
Object.excel_data()
Total_count = len(Object.xl_Event_id)
print "Number of Rows ::", Total_count

if Object.login == 'OK':
    for looping in range(0, Total_count):
        print "Iteration Count is ::", looping
        Object.schedule_interview(looping)
        Object.provide_feedback(looping)
