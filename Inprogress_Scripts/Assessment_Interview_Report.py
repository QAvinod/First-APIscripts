import requests
import json
import time
import urllib2


class EventAssessmentInterviewReport:
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
        self.get_Activation_token = {"content-type": "application/json",
                                     "X-AUTH-TOKEN": self.response.get("Token")}
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

    def download_report(self):
        self.report_request = {
            "EventId": 1065,
            "JobId": 842,
            "Sync": "True",
            "AddInterviewDetails": "True",
            "AddSkillDetails": "False",
            "AddTestsDetails": "True",
            "StatusIds": [72277, 72329, 72331, 72332, 72333, 72334, 72335, 72699, 72700, 72337, 72338, 72339, 72340,
                          72341, 72360, 72343, 72344, 72593, 72355, 72356, 72358, 72359, 72363, 72579, 72367, 72368,
                          72369, 72370, 72371, 72543, 72545, 72546, 72547, 72548, 72549, 72718, 72719, 72551, 72553,
                          72554, 72555, 72556, 72558, 72604, 72605, 72606, 72607, 72608, 72609, 72709, 72710, 72711],
            "Columns": {
                "CandidateProperties": [{
                    "catalog_id": 29297,
                    "ColumnName": "Candidate Id"
                }, {
                    "catalog_id": 29301,
                    "ColumnName": "Full Name"
                }],
                "TaskProperties": [{
                    "ColumnName": "gender",
                    "ColumnHeader": "Gender"
                }, {
                    "ColumnName": "preferredWorkLocation1",
                    "ColumnHeader": "Preferred Work Location 1"
                }]
            }
        }
        report_api = requests.post("https://amsin.hirepro.in/py/crpo/api/v1/applicantExportPtyNew1",
                                   headers=self.get_Activation_token,
                                   data=json.dumps(self.report_request, default=str), verify=False)
        report_api_dict = json.loads(report_api.content)
        status = report_api_dict['status']
        data = report_api_dict['data']
        link = json.loads(data)
        download_link = link['downloadLink']
        # print download_link

        # ----------------------
        # Download Excel Report
        # ----------------------
        link_path = download_link
        resp = requests.get(link_path)
        with open('/home/vinod/HireproApp/AIR_OUTPUT/Assessment_interview_Report.xlsx', 'wb') as output:
            output.write(resp.content)

        if status == 'OK':
            print 'Successfully generated Link'
        else:
            print 'API has been failed to generate Link'


Object = EventAssessmentInterviewReport()
if Object.status == 'OK':
    Object.download_report()
