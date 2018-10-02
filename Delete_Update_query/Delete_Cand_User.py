import datetime
import mysql
import time
import xlrd
from mysql import connector


class DeleteQuery:
    def __init__(self):

        self.now = datetime.datetime.now()
        self.password = raw_input('DB password ::')

        # --------------------------
        # Initialising Excel Data
        # --------------------------
        self.xl_candidateId = []
        self.xl_userId = []
        self.xl_testuserId = []

    def candidate_excel_data(self):

        # ---------------------------
        # CandidateId Excel Data Read
        # ---------------------------
        workbook = xlrd.open_workbook('/home/vinod/SOFTWARE/InputFiles/UploadCandidate/OutPut/API_UploadCandidates.xls')
        sheet1 = workbook.sheet_by_index(0)
        for i in range(1, sheet1.nrows):
            number = i  # Counting number of rows
            rows = sheet1.row_values(number)

            if not rows[2]:
                self.xl_candidateId.append(None)
            else:
                self.xl_candidateId.append(int(rows[2]))

    def user_excel_data(self):

        # ----------------
        # Excel Data Read
        # ----------------
        workbook = xlrd.open_workbook('/home/vinod/SOFTWARE/InputFiles/User/Output/API_Create_User.xls')
        sheet1 = workbook.sheet_by_index(0)
        for i in range(1, sheet1.nrows):
            number = i  # Counting number of rows
            rows = sheet1.row_values(number)

            if not rows[2]:
                self.xl_userId.append(None)
            else:
                self.xl_userId.append(int(rows[2]))

    def test_user_excel_data(self):

        # ----------------
        # Excel Data Read
        # ----------------
        workbook = xlrd.open_workbook('/home/vinod/SOFTWARE/InputFiles/ScoreSheet/UploadScores.xls')
        sheet1 = workbook.sheet_by_index(1)
        for i in range(1, sheet1.nrows):
            number = i  # Counting number of rows
            rows = sheet1.row_values(number)

            if not rows[0]:
                self.xl_testuserId.append(None)
            else:
                self.xl_testuserId.append(int(rows[0]))

    def dbconnection(self):

        # -------------
        # DB Connection
        # -------------
        self.connection = mysql.connector.connect(host='35.154.36.218',
                                                  database='appserver_core',
                                                  user='qauser',
                                                  password=self.password)
        self.cursor = self.connection.cursor()

    def deleteusers(self, loop):
        self.dbconnection()

        # -------------------
        # User archive query
        # -------------------
        if self.xl_userId[loop]:
            self.user_query = "UPDATE appserver_core.users SET tenant_id='0', is_archived='1'," \
                              " is_deleted='1' WHERE id = %s" % self.xl_userId[loop]
            print self.user_query
            query = self.user_query
            time.sleep(2)
            self.cursor.execute(query)
            self.connection.commit()

    def deletecandidates(self, loop):
        self.dbconnection()

        # -------------------------------
        # Candidate delete & update query
        # -------------------------------
        if self.xl_candidateId[loop]:
            self.candidate_query = "UPDATE appserver_core.candidates SET tenant_id=0, is_archived=1," \
                                   "is_deleted=1, is_draft=0 WHERE id = %s" % (self.xl_candidateId[loop])
            print self.candidate_query
            query = self.candidate_query
            time.sleep(2)
            self.cursor.execute(query)
            self.connection.commit()

            time.sleep(2)
            self.candidate_query1 = "DELETE FROM duplicate_candidates_infos where candidate_id=%s;" \
                                    % (self.xl_candidateId[loop])
            print self.candidate_query1
            query1 = self.candidate_query1
            time.sleep(2)
            self.cursor.execute(query1)
            self.connection.commit()

            time.sleep(2)
            self.candidate_query2 = "DELETE FROM appserver_core.candidates WHERE tenant_id=0 and id =%s;"\
                                    % (self.xl_candidateId[loop])
            print self.candidate_query2
            query2 = self.candidate_query2
            time.sleep(2)
            self.cursor.execute(query2)
            self.connection.commit()

            time.sleep(2)
            self.candidate_query3 = "DELETE FROM test_users WHERE id =%s;" % (self.xl_candidateId[loop])
            print self.candidate_query3
            query2 = self.candidate_query3
            time.sleep(2)
            self.cursor.execute(query2)
            self.connection.commit()

    def delete_testuser_score(self, loop):
        self.dbconnection()

        # ------------
        # Score delete
        # ------------
        self.testuser_query = "delete from candidate_scores where testuser_id =%s;" % (self.xl_testuserId[loop])
        print self.testuser_query
        query = self.testuser_query
        time.sleep(2)
        self.cursor.execute(query)
        self.connection.commit()

        # ------------
        # update query
        # ------------
        self.testuser_query1 = "UPDATE test_users SET total_score ='', percentage ='', status='' WHERE id=%s;" \
                               % (self.xl_testuserId[loop])
        print self.testuser_query1
        query = self.testuser_query1
        time.sleep(2)
        self.cursor.execute(query)
        self.connection.commit()


Object = DeleteQuery()
Object.candidate_excel_data()
Object.user_excel_data()
Object.test_user_excel_data()

Total_count = len(Object.xl_candidateId)
if Object.xl_candidateId:
    for looping in range(0, Total_count):
        Object.deletecandidates(looping)

Total_count_1 = len(Object.xl_userId)
if Object.xl_userId:
    for looping in range(0, Total_count_1):
        Object.deleteusers(looping)

Total_count_1 = len(Object.xl_testuserId)
if Object.xl_testuserId:
    for looping in range(0, Total_count_1):
        Object.delete_testuser_score(looping)




