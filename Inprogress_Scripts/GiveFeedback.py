import time
import json
import requests
import xlwt
import xlrd
import datetime
import exceptions


class InterviewFeedback:
    def __init__(self):
        # ------------------------
        # CRPO LOGIN APPLICATION
        # ------------------------
        try:
            self.header = {"content-type": "application/json"}
            # self.TenantAlias = raw_input('TenantAlias:: ')
            # self.LoginName = raw_input('LoginName:: ')
            # self.Password = raw_input('Password:: ')
            self.login_request = {"LoginName": "admin",
                                  "Password": "Mohi@12345",
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
        except exceptions.ValueError as login_error:
            print(login_error)
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
        self.__style9 = xlwt.easyxf('pattern: pattern solid, fore_colour light_green;'
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
        excelheaders = ['Comparison', 'Status', 'Schedule_Message', 'IR_id', 'Feedback_Message', 'Interviewer_Decision',
                        'Partial_Feedback', 'Partial_Feedback_message', 'Scheduled_date', 'Interviewed_date', 'Skill_01',
                        'Score_01', 'Skill_02', 'Score_02', 'Skill_03', 'Score_03', 'Skill_04', 'Score_04', 'Duration',
                        'Skill_comment', 'OverAllComment', 'Update_Duration', 'Update_Skill_comment',
                        'Updated_OverAllComment']
        for headers in excelheaders:
            if headers in ['Comparison', 'Status', 'Schedule_Message', 'IR_id', 'Feedback_Message',
                           'Interviewer_Decision', 'Partial_Feedback', 'Partial_Feedback_message']:
                self.ws.write(0, index, headers, self.__style2)
            elif headers in ['Scheduled_date', 'Interviewed_date', 'Update_Skill_comment', 'Updated_OverAllComment',
                             'Update_Duration']:
                self.ws.write(0, index, headers, self.__style9)
            else:
                self.ws.write(0, index, headers, self.__style0)
            index += 1
        print ('Excel Headers are printed successfully')
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
        self.xl_Schedule_Comment = []
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
        self.xl_Over_all_comment = []
        self.xl_partial_feedback = []
        # -----------------------------------------
        # Update details / Partial feedback details
        # -----------------------------------------
        self.xl_updated_duration = []
        self.xl_Updated_Over_all_comment = []
        self.xl_update_Skill_comment = []
        # -----------------------------------------------------------------------------------
        # Dictionaries for Interview_schedule, interview_feedback, interview_feedback_details
        # -----------------------------------------------------------------------------------
        self.ir = {}
        self.i_r = self.ir
        self.is_success = {}
        self.is_s = self.is_success
        self.is_feedback = {}
        self.i_f = self.is_feedback
        self.message = {}
        self.m = self.message
        self.feedback = {}
        self.f = self.feedback
        self.feedback_data = {}
        self.fd = self.feedback_data
        self.updated_feedback_data = {}
        self.u_fd = self.updated_feedback_data
        # -------------------
        # Skill dictionaries
        # -------------------
        self.skill_dict_1 = {}
        self.skill_1 = self.skill_dict_1
        self.skill_dict_2 = {}
        self.skill_2 = self.skill_dict_2
        self.skill_dict_3 = {}
        self.skill_3 = self.skill_dict_3
        self.skill_dict_4 = {}
        self.skill_4 = self.skill_dict_4
        self.filledFeedbackDetails = {}
        self.ffd = self.filledFeedbackDetails
        self.skillAssessed_details = {}
        self.sad = self.skillAssessed_details
        # ---------------------------
        # Skill updated dictionaries
        # ---------------------------
        self.updated_skill_dict_1 = {}
        self.updated_skill_1 = self.updated_skill_dict_1
        self.updated_skill_dict_2 = {}
        self.updated_skill_2 = self.updated_skill_dict_2
        self.updated_skill_dict_3 = {}
        self.updated_skill_3 = self.updated_skill_dict_3
        self.updated_skill_dict_4 = {}
        self.updated_skill_4 = self.updated_skill_dict_4
        self.updated_filledFeedbackDetails = {}
        self.updated_ffd = self.updated_filledFeedbackDetails
        self.updated_skillAssessed_details = {}
        self.u_sad = self.updated_skillAssessed_details
        # ----------------------------
        # Partial/updated Dictionaries
        # ----------------------------
        self.pf = {}
        self.p_f = self.pf

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
                    self.xl_Schedule_Comment.append(str(rows[7]))
                else:
                    self.xl_Schedule_Comment.append(None)
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
                if rows[21] is not None and rows[21] != '':
                    self.xl_Over_all_comment.append(str(rows[21]))
                else:
                    self.xl_Over_all_comment.append(None)
                if rows[22] is not None and rows[22] != '':
                    self.xl_partial_feedback.append(int(rows[22]))
                else:
                    self.xl_partial_feedback.append(None)
                if rows[23] is not None and rows[23] != '':
                    self.xl_updated_duration.append(int(rows[23]))
                else:
                    self.xl_updated_duration.append(None)
                if rows[24] is not None and rows[24] != '':
                    self.xl_Updated_Over_all_comment.append(str(rows[24]))
                else:
                    self.xl_Updated_Over_all_comment.append(None)
                if rows[25] is not None and rows[25] != '':
                    self.xl_update_Skill_comment.append(str(rows[25]))
                else:
                    self.xl_update_Skill_comment.append(None)
            print('Excel data initiated is Done')
        except IOError:
            print("File not found or path is incorrect")

    def schedule_interview(self, loop):
        try:
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
                "recruiterComment": self.xl_Schedule_Comment[loop],
                "recruitEventId": self.xl_Event_id[loop],
                "applicantIds": [self.xl_Applicant_id[loop]]
            }]
            scheduling_interviews = requests.post("https://amsin.hirepro.in/py/crpo/api/v1/interview/schedule/",
                                                  headers=self.get_token,
                                                  data=json.dumps(schedule_request, default=str), verify=False)
            schedule_response = json.loads(scheduling_interviews.content)
            # print (json.dumps(schedule_response, indent=2))
            data = schedule_response['data']
            # print(json.dumps(data, indent=2))
            # print('***--------------------------------------------------------***')
            if schedule_response['status'] == 'OK':
                success = data['success']
                failure = data['failure']
                if data['success']:
                    for i in success:
                        self.ir = i['interviewRequestId']
                        print self.ir
                        print "Scheduled to interview"
                        self.message = i.get('message')
                        self.is_success = True
                elif data['failure']:
                    for i in failure:
                        self.message = i.get('message')
                        print self.message
                        self.is_success = False
            else:
                print ('Error occured while scheduling')
        except exceptions.ValueError as Schedule_error:
            print(Schedule_error)

    def provide_feedback(self, loop):
        if self.xl_int_datetime[loop]:
            if self.xl_partial_feedback[loop] == 1:
                self.pf = True
            else:
                self.pf = False
            try:
                feedback_request = {
                    "interviewRequestId": self.ir,
                    "interviewerFeedback": [{
                        "partial_feedback": self.pf,
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
                providing_feedback = requests.post("https://amsin.hirepro.in/py/crpo/api/v1/interview/givefeedback/",
                                                   headers=self.get_token,
                                                   data=json.dumps(feedback_request, default=str), verify=False)
                feedback_response = json.loads(providing_feedback.content)
                # print (json.dumps(feedback_response, indent=2))
                data = feedback_response['data']
                self.feedback_message = data['message']
                self.is_feedback = True
                print('Provide Feedback is Done')
            except exceptions.ValueError as feedback_error:
                print(feedback_error)

    def feedback_details(self, loop):
        try:
            details_url = requests.get("https://amsin.hirepro.in/py/crpo/api/v1/interview/get/{}".format(self.ir),
                                       headers=self.get_token)
            details_response = json.loads(details_url.content)
            # print(json.dumps(details_response, indent=2))
            # print('***--------------------------------------------------------***')
            self.feedback_data = details_response['data']
            self.filledFeedbackDetails = self.feedback_data['filledFeedbackDetails']
            for feedback in self.filledFeedbackDetails:
                self.feedback = feedback
                for skillAssessed_details in feedback['skillAssessed']:
                    self.skillAssessed_details = skillAssessed_details
                    if self.xl_Skill_id_01[loop] == skillAssessed_details['skillId']:
                        self.skill_dict_1 = skillAssessed_details
                    if self.xl_Skill_id_02[loop] == skillAssessed_details['skillId']:
                        self.skill_dict_2 = skillAssessed_details
                    if self.xl_Skill_id_03[loop] == skillAssessed_details['skillId']:
                        self.skill_dict_3 = skillAssessed_details
                    if self.xl_Skill_id_04[loop] == skillAssessed_details['skillId']:
                        self.skill_dict_4 = skillAssessed_details
            print('Feedback details are fetched Successfully')
        except exceptions.ValueError as details_error:
            print(details_error)

    def updated_feedback_details(self, loop):
        try:
            details_url = requests.get("https://amsin.hirepro.in/py/crpo/api/v1/interview/get/{}".format(self.ir),
                                       headers=self.get_token)
            details_response = json.loads(details_url.content)
            # print(json.dumps(details_response, indent=2))
            # print('***--------------------------------------------------------***')
            self.updated_feedback_data = details_response['data']
            self.updated_filledFeedbackDetails = self.updated_feedback_data['filledFeedbackDetails']
            for updated_feedback in self.updated_filledFeedbackDetails:
                for updated_skillAssessed_details in updated_feedback['skillAssessed']:
                    self.updated_skillAssessed_details = updated_skillAssessed_details
                    if self.xl_Skill_id_01[loop] == updated_skillAssessed_details['skillId']:
                        self.updated_skill_dict_1 = updated_skillAssessed_details
                    if self.xl_Skill_id_02[loop] == updated_skillAssessed_details['skillId']:
                        self.updated_skill_dict_2 = updated_skillAssessed_details
                    if self.xl_Skill_id_03[loop] == updated_skillAssessed_details['skillId']:
                        self.updated_skill_dict_3 = updated_skillAssessed_details
                    if self.xl_Skill_id_04[loop] == updated_skillAssessed_details['skillId']:
                        self.updated_skill_dict_4 = updated_skillAssessed_details
            print('Updated Feedback details are fetched Successfully')
        except exceptions.ValueError as details_error:
            print(details_error)

    def partial_feedback(self, loop):
        if self.feedback['partialFeedback'] == 1:
            try:
                update_feedback = {
                    "InterviewRequestId": self.ir,
                    "FilledFormId": self.skillAssessed_details['interviewfilledfeedbackformId'],
                    "Duration": self.xl_updated_duration[loop],
                    "OverAllComments": self.xl_Updated_Over_all_comment[loop],
                    "Skills": [{
                        "Id": self.skill_dict_1['id'],
                        "SkillId": self.xl_Skill_id_01[loop],
                        "Comments": self.xl_update_Skill_comment[loop],
                        "SkillRating": self.xl_Skill_score_01[loop]
                    }, {
                        "Id": self.skill_dict_2['id'],
                        "SkillId": self.xl_Skill_id_02[loop],
                        "Comments": self.xl_update_Skill_comment[loop],
                        "SkillRating": self.xl_Skill_score_02[loop]
                    }, {
                        "Id": self.skill_dict_3['id'],
                        "SkillId": self.xl_Skill_id_03[loop],
                        "Comments": self.xl_update_Skill_comment[loop],
                        "SkillRating": self.xl_Skill_score_03[loop]
                    }, {
                        "Id": self.skill_dict_4['id'],
                        "SkillId": self.xl_Skill_id_04[loop],
                        "Comments": self.xl_update_Skill_comment[loop],
                        "SkillRating": self.xl_Skill_score_04[loop]
                    }]
                }
                partial_url = requests.post(
                    "https://amsin.hirepro.in/py/crpo/api/v1/interview/updateinterviewerfeedback",
                    headers=self.get_token,
                    data=json.dumps(update_feedback, default=str), verify=False)
                partial_response = json.loads(partial_url.content)
                self.partial_data = partial_response['data']
            except exceptions.ValueError as Partial_update_error:
                print(Partial_update_error)

    def output_excel(self, loop):
        # ------------------
        # Writing Input Data
        # ------------------
        self.ws.write(self.rowsize, self.col, 'Input', self.__style4)
        self.ws.write(self.rowsize, 8, self.xl_Datetime[loop], self.__style1)
        self.ws.write(self.rowsize, 9, self.xl_int_datetime[loop], self.__style1)
        self.ws.write(self.rowsize, 10, self.xl_Skill_id_01[loop], self.__style1)
        self.ws.write(self.rowsize, 11, self.xl_Skill_score_01[loop], self.__style1)
        self.ws.write(self.rowsize, 12, self.xl_Skill_id_02[loop], self.__style1)
        self.ws.write(self.rowsize, 13, self.xl_Skill_score_02[loop], self.__style1)
        self.ws.write(self.rowsize, 14, self.xl_Skill_id_03[loop], self.__style1)
        self.ws.write(self.rowsize, 15, self.xl_Skill_score_03[loop], self.__style1)
        self.ws.write(self.rowsize, 16, self.xl_Skill_id_04[loop], self.__style1)
        self.ws.write(self.rowsize, 17, self.xl_Skill_score_04[loop], self.__style1)
        self.ws.write(self.rowsize, 18, self.xl_duration[loop], self.__style1)
        self.ws.write(self.rowsize, 19, self.xl_skill_comment[loop], self.__style1)
        self.ws.write(self.rowsize, 20, self.xl_Over_all_comment[loop], self.__style1)
        self.ws.write(self.rowsize, 21, self.xl_updated_duration[loop], self.__style1)
        self.ws.write(self.rowsize, 22, self.xl_update_Skill_comment[loop], self.__style1)
        self.ws.write(self.rowsize, 23, self.xl_Updated_Over_all_comment[loop], self.__style1)
        # -------------------
        # Writing Output data
        # -------------------
        self.rowsize += 1  # Row increment
        self.ws.write(self.rowsize, self.col, 'Output', self.__style5)
        # --------------------------------------------------------------------------------------------------------------
        self.ws.write(self.rowsize, 1, 'status', self.__style8)
        # --------------------------------------------------------------------------------------------------------------
        if self.is_success:
            self.ws.write(self.rowsize, 2, self.feedback_data['interviewerComment'], self.__style8)
        else:
            self.ws.write(self.rowsize, 2, self.message, self.__style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.ir:
            self.ws.write(self.rowsize, 3, self.ir, self.__style8)
        else:
            self.ws.write(self.rowsize, 3, None)
        # --------------------------------------------------------------------------------------------------------------
        if self.is_feedback:
            self.ws.write(self.rowsize, 4, self.feedback_message, self.__style8)
        else:
            self.ws.write(self.rowsize, 4, 'Error occured while giving feedback', self.__style3)
        # --------------------------------------------------------------------------------------------------------------
        if self.feedback and self.feedback['decisionText']:
            self.ws.write(self.rowsize, 5, self.feedback['decisionText'], self.__style8)
        else:
            self.ws.write(self.rowsize, 5, None)
        # --------------------------------------------------------------------------------------------------------------
        if self.feedback and self.feedback['partialFeedback'] == 1:
            self.ws.write(self.rowsize, 6, 'True', self.__style8)
        elif self.feedback and self.feedback['partialFeedback'] == 0:
            self.ws.write(self.rowsize, 6, 'False', self.__style3)
        # -------------------------------------------------------------------------------------------------------------
        if self.partial_data.get('message'):
            self.ws.write(self.rowsize, 7, self.partial_data['message'], self.__style8)
        else:
            self.ws.write(self.rowsize, 7, None)
        # --------------------------------------------------------------------------------------------------------------
        if self.xl_Datetime == self.feedback_data.get('interviewTime'):
            self.ws.write(self.rowsize, 8, self.feedback_data.get('interviewTime'), self.__style8)
        else:
            self.ws.write(self.rowsize, 8, self.feedback_data.get('interviewTime', 'NA'), self.__style3)
        # --------------------------------------------------------------------------------------------------------------
        # if self.filledFeedbackDetails == self.filledFeedbackDetails.get('interviewedTime'):
        #     self.ws.write(self.rowsize, 9, self.filledFeedbackDetails.get('interviewedTime'), self.__style8)
        # else:
        #     self.ws.write(self.rowsize, 9, self.filledFeedbackDetails.get('interviewedTime', 'NA'), self.__style3)
        # # --------------------------------------------------------------------------------------------------------------
        #
        # if self.filledFeedbackDetails and self.filledFeedbackDetails.get('duration'):
        #     self.ws.write(self.rowsize, 18, self.filledFeedbackDetails.get('duration'))
        # else:
        #     self.ws.write(self.rowsize, 18, self.filledFeedbackDetails.get('duration', 'NA'), self.__style3)
        # # --------------------------------------------------------------------------------------------------------------
        #
        # if self.updated_filledFeedbackDetails and self.updated_filledFeedbackDetails.get('duration'):
        #     self.ws.write(self.rowsize, 21, self.updated_filledFeedbackDetails.get('duration'))
        # else:
        #     self.ws.write(self.rowsize, 21, self.updated_filledFeedbackDetails.get('duration', 'NA'), self.__style3)
        # # --------------------------------------------------------------------------------------------------------------
        self.rowsize += 1  # Row increment
        Object.wb_Result.save('/home/vinod/Desktop/Output/API_GiveFeedback.xls')
        print('Excel data is ready')


Object = InterviewFeedback()
Object.excel_data()
Total_count = len(Object.xl_Event_id)
print "Number of Rows ::", Total_count
try:
    if Object.login == 'OK':
        for looping in range(0, Total_count):
            print "Iteration Count is ::", looping
            Object.schedule_interview(looping)
            if Object.is_success:
                Object.provide_feedback(looping)
                Object.feedback_details(looping)
                if Object.pf:
                    Object.partial_feedback(looping)
                    Object.updated_feedback_details(looping)
            Object.output_excel(looping)
            # -------------------------------------
            # Making all dict empty for every loop
            # -------------------------------------
            Object.is_success = {}
            Object.message = {}
            Object.ir = {}
            Object.feedback_data = {}
            Object.feedback = {}
            Object.is_feedback = {}
            Object.updated_feedback_data = {}
            # ----------
            # Skill dict
            # ----------
            Object.skill_dict_1 = {}
            Object.skill_dict_2 = {}
            Object.skill_dict_3 = {}
            Object.skill_dict_4 = {}
            Object.filledFeedbackDetails = {}
            Object.skillAssessed_details = {}
            # ------------------
            # updated Skill dict
            # ------------------
            Object.updated_skill_dict_1 = {}
            Object.updated_skill_dict_2 = {}
            Object.updated_skill_dict_3 = {}
            Object.updated_skill_dict_4 = {}
            Object.updated_filledFeedbackDetails = {}
            Object.updated_skillAssessed_details = {}
            # ---------------------
            # Partial/updated  dict
            # ---------------------
            Object.pf = {}
except exceptions.AttributeError as Object_error:
    print(Object_error)