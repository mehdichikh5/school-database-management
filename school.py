from mail import *
import smtplib

from PyQt5.QtCore import *

from PyQt5.QtGui import *

from PyQt5.QtWidgets import *

import mysql.connector as con

from PyQt5.uic import loadUiType

import sys

from PyQt5.QtPrintSupport import QPrinter, QPrintDialog,QPrintPreviewDialog

from datetime import date

import pandas as pd



ui, _ = loadUiType('school.ui')



class MainApp(QMainWindow, ui):

    def __init__(self):

        QMainWindow.__init__(self)

        self.setupUi(self)


        self.tabWidget.setCurrentIndex(0)

        self.menubar.setVisible(False)

        self.tabWidget.tabBar().setVisible(False)

        self.b01.clicked.connect(self.login)


        self.b11.clicked.connect(self.add_new_student)

        self.menu11.triggered.connect(self.show_add_new_student)

        self.menu12.triggered.connect(self.show_add_edit_student)

        self.menu13.triggered.connect(self.show_report)

        self.menu21.triggered.connect(self.show_mark)

        self.menu31.triggered.connect(self.show_attendance)

        self.menu51.triggered.connect(self.show_fees)

        self.menu61.triggered.connect(self.show_report)

        self.menu62.triggered.connect(self.show_report)

        self.menu63.triggered.connect(self.show_report)

        self.menu64.triggered.connect(self.show_report)

        self.menu71.triggered.connect(self.logout)

       

        self.cb21.currentIndexChanged.connect(self.fill_details_on_combobox_selected)

        self.b21.clicked.connect(self.update_student)

        self.b22.clicked.connect(self.delete_student)


        self.cb32.currentIndexChanged.connect(self.fill_examname_on_combobox_selected)

        self.b32.clicked.connect(self.fill_marks_on_button_getmarks_clicked)

        self.b31.clicked.connect(self.save_mark)

        self.b33.clicked.connect(self.update_mark)

        self.b34.clicked.connect(self.delete_mark)


        self.b41.clicked.connect(self.save_attendance)

        self.cb42.currentIndexChanged.connect(self.fill_date_on_combobox_selected)

        self.b42.clicked.connect(self.fill_attendance_on_button_getattendance_clicked)

        self.b44.clicked.connect(self.delete_attendance)

        self.b43.clicked.connect(self.update_attendance)


        self.b51.clicked.connect(self.save_fees)

        self.cb52.currentIndexChanged.connect(self.fill_fees_details_on_receipt_number_selected)

        self.b52.clicked.connect(self.update_fees)

        self.b53.clicked.connect(self.delete_fees)


        self.button_export.clicked.connect(self.export_table)


        self.b71.clicked.connect(self.print_file)

        self.b72.clicked.connect(self.cancel_print)


    ######## ADMIN LOGIN FORM ##########


    def login(self):

        un = self.tb01.text()

        pw = self.tb02.text()

        if (un == "admin" and pw == "admin"):

            self.menubar.setVisible(True)

            self.tabWidget.setCurrentIndex(1)

        else:

            QMessageBox.information(self, "School management system", "Invalid login details, Try again !")

            self.l01.setText("Invalid login details, Try again !")


    ######## NEW STUDENT CREATION FORM ##########


    def show_add_new_student(self):

        self.tabWidget.setCurrentIndex(2)

        self.fill_next_registration_number()


    def fill_next_registration_number(self):

        try:

            rn=0

            mydb = con.connect(host="localhost",port = 3307, user="root", password="motdepasse", db="school")


            cursor = mydb.cursor()

            cursor.execute("SELECT * FROM student")

            result = cursor.fetchall()

            if result:

                for stud in result:

                    rn += 1

            self.tb11.setText(str(rn+1))

        except con.Error as e:

            print("Select error" + str(e))


    def add_new_student(self):

        try:
            mydb = con.connect(host="localhost",port = 3307, user="root", password="motdepasse", db="school")
            
            cursor = mydb.cursor()

            
            
            registration_number = self.tb11.text()

            full_name = self.tb12.text()

            gender = self.cb11.currentText()

            date_of_birth = self.tb13.text()

            age = self.tb14.text()

            adress = self.mtb11.toPlainText()

            phone = self.tb15.text()

            email = self.tb16.text()

            standard = self.cb12.currentText()

            value = (registration_number, full_name, gender, date_of_birth, age, adress, phone, email, standard)


            cursor.execute("INSERT INTO student (registration_number,full_name,gender,date_of_birth,age,adress,phone,email,standard) VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s)",value)


            
            mydb.commit()

            self.l11.setText("Student Details Saved successfully!")

            QMessageBox.information(self, "School Management System", "Student Details Saved successfully!")

            self.tabWidget.setCurrentIndex(1)

        except con.Error as e:

            self.l11.setText("Error in student creation! " + str(e))

    ######## EDIT AND DELETE STUDENT FORM ##########


    def show_add_edit_student(self):

        self.tabWidget.setCurrentIndex(3)

        self.fill_student_regno_combobox()


    def fill_student_regno_combobox(self):

        try:

            mydb = con.connect(host="localhost",port = 3307, user="root", password="", db="school")

            cursor = mydb.cursor()

            cursor.execute("SELECT * FROM student")

            self.cb21.clear()

            result = cursor.fetchall()

            if result:

                for stud in result:

                    # self.c21.addItem(stud[0])

                    # print(stud[0])

                    self.cb21.addItem(str(stud[1]))

        except con.Error as e:

            print("Select error" + e)


    def fill_details_on_combobox_selected(self):

        try:

            mydb = con.connect(host="localhost",port = 3307, user="root", password="", db="school")

            cursor = mydb.cursor()

            cursor.execute("SELECT * FROM student where registration_number = '" + self.cb21.currentText() + "'")

            result = cursor.fetchall()

            if result:

                for stud in result:

                    self.tb21.setText(str(stud[2]))

                    self.tb22.setText(str(stud[3]))

                    self.tb23.setText(str(stud[4]))

                    self.tb24.setText(str(stud[5]))

                    self.mtb21.setText(str(stud[6]))

                    self.tb25.setText(str(stud[7]))

                    self.tb26.setText(str(stud[8]))

                    self.tb27.setText(str(stud[9]))

        except con.Error as e:
            print(e)
            print("Select error" + e)


    def update_student(self):

        try:

            mydb = con.connect(host="localhost",port = 3307, user="root", password="", db="school")

            cursor = mydb.cursor()

            registration_number = self.cb21.currentText()

            full_name = self.tb21.text()

            gender = self.tb22.text()

            date_of_birth = self.tb23.text()

            age = self.tb24.text()

            adress = self.mtb21.toPlainText()

            phone = self.tb25.text()

            email = self.tb26.text()

            standard = self.tb27.text()

            qry = "UPDATE student set full_name = '" + full_name + "',gender= '" + gender + "',date_of_birth= '" + date_of_birth + "',age= '" + age + "',adress= '" + adress + "',phone= '" + phone + "',email= '" + email + "',standard= '" + standard + "' where registration_number ='" + registration_number + "' "

            cursor.execute(qry)

            mydb.commit()

            self.l21.setText("Student Details Modified successfully!")

            QMessageBox.information(self, "School Management System", "Student Details Modified successfully!")

            self.tabWidget.setCurrentIndex(1)

        except con.Error as e:

            self.l21.setText("Error in student creation!")


    def delete_student(self):

        m=QMessageBox.question(self,"Delete","Are you sure want to delete",QMessageBox.Yes|QMessageBox.No)

        if m == QMessageBox.Yes:

            try:

                mydb = con.connect(host="localhost",port = 3307, user="root", password="", db="school")

                cursor = mydb.cursor()

                registration_number = self.cb21.currentText()

                qry = "DELETE from student where registration_number ='" + registration_number + "' "

                cursor.execute(qry)

                mydb.commit()

                print("Deleted")

                self.fill_student_regno_combobox()

                self.l21.setText("Student Details Deleted successfully!")

                QMessageBox.information(self, "School Management System", "Student Details Deleted successfully!")

                self.tabWidget.setCurrentIndex(1)

            except con.Error as e:

                self.l21.setText("Error in student deletion!")

        else:

            print("Not")



    ######## MARK DETAILS ##########


    def show_mark(self):

        self.tabWidget.setCurrentIndex(4)

        self.fill_student_regno_combobox_mark()


    def fill_student_regno_combobox_mark(self):

        try:

            mydb = con.connect(host="localhost",port = 3307, user="root", password="", db="school")

            cursor = mydb.cursor()

            cursor.execute("SELECT * FROM student")

            self.cb31.clear()

            self.cb32.clear()

            result = cursor.fetchall()

            if result:

                for stud in result:

                    # self.c21.addItem(stud[0])

                    # print(stud[0])

                    self.cb31.addItem(str(stud[1]))

                    self.cb32.addItem(str(stud[1]))

        except con.Error as e:

            print("Select error" + e)


    def save_mark(self):

        try:

            mydb = con.connect(host="localhost",port = 3307, user="root", password="", db="school")

            cursor = mydb.cursor()

            registration_number = self.cb31.currentText()

            exam_name = self.tb31.text()

            grade = self.tb32.text()

            
            qry = "INSERT INTO mark (registration_number,exam_name,grade) VALUES(%s,%s,%s)"

            value = (registration_number, exam_name, grade)

            cursor.execute(qry, value)

            mydb.commit()

            self.l11.setText("Mark Details Saved successfully!")

            QMessageBox.information(self, "School Management System", "Mark Details Saved successfully!")

            self.tabWidget.setCurrentIndex(1)

        except con.Error as e:
            print(e)
            self.l11.setText("Error in mark details save !")


    def fill_examname_on_combobox_selected(self):

        try:

            self.cb33.clear()

            mydb = con.connect(host="localhost",port = 3307, user="root", password="", db="school")

            cursor = mydb.cursor()

            cursor.execute("SELECT * FROM mark where registration_number = '" + self.cb32.currentText() + "'")

            print("SELECT * FROM mark where registration_number = '" + self.cb32.currentText() + "'")

            result = cursor.fetchall()

            print(result)

            if result:

                for exam in result:

                    self.cb33.addItem(str(exam[2]))

        except con.Error as e:

            print("Select error" + e)


    def fill_marks_on_button_getmarks_clicked(self):

        try:

            mydb = con.connect(host="localhost",port = 3307, user="root", password="", db="school")

            cursor = mydb.cursor()

            cursor.execute(

                "SELECT * FROM mark where registration_number = '" + self.cb32.currentText() + "' and exam_name = '" + self.cb33.currentText() + "'")

            print("SELECT * FROM mark where registration_number = '" + self.cb32.currentText() + "'")

            result = cursor.fetchall()

            if result:

                for stud in result:

                    self.tb37.setText(str(stud[3]))



        except con.Error as e:

            print("Select error" + e)


    def update_mark(self):

        try:

            mydb = con.connect(host="localhost",port = 3307, user="root", password="", db="school")

            cursor = mydb.cursor()

            registration_number = self.cb32.currentText()

            exam_name = self.cb33.currentText()

            grade = self.tb37.text()

            

            qry = "UPDATE mark set registration_number = '" + registration_number + "',exam_name= '" + exam_name + "',grade = '" + grade + "'"

            print(qry)

            cursor.execute(qry)

            self.tb37.setText("")

            mydb.commit()

            self.l32.setText("Mark Details Modified successfully!")

            QMessageBox.information(self, "School Management System", "Mark Details Modified successfully!")

            self.tabWidget.setCurrentIndex(1)

        except con.Error as e:
            print(e)
            self.l32.setText("Error in student modification!")


    def delete_mark(self):

        m=QMessageBox.question(self,"Delete","Are you sure want to delete",QMessageBox.Yes|QMessageBox.No)

        if m == QMessageBox.Yes:

            try:

                mydb = con.connect(host="localhost",port = 3307, user="root", password="", db="school")

                cursor = mydb.cursor()

                registration_number = self.cb32.currentText()

                exam_name = self.cb33.currentText()

                qry = "DELETE from mark where registration_number ='" + registration_number + "' and exam_name = '" + exam_name + "' "

                cursor.execute(qry)

                mydb.commit()

                self.tb37.setText("")


                self.fill_student_regno_combobox()

                self.l32.setText("Mark Details Deleted successfully!")

                QMessageBox.information(self, "School Management System", "Mark Details Deleted successfully!")

                self.tabWidget.setCurrentIndex(1)

            except con.Error as e:

                self.l32.setText("Error in mark deletion!")


    ######## ATTENDANCE DETAILS ##########


    def show_attendance(self):

        self.tabWidget.setCurrentIndex(5)

        self.tb41.setText(str(date.today()))

        self.fill_student_regno_combobox_attendance()


    def fill_student_regno_combobox_attendance(self):

        try:

            mydb = con.connect(host="localhost",port = 3307, user="root", password="", db="school")

            cursor = mydb.cursor()

            cursor.execute("SELECT * FROM student")

            self.cb41.clear()

            self.cb42.clear()

            result = cursor.fetchall()

            if result:

                for stud in result:

                    self.cb41.addItem(str(stud[1]))

                    self.cb42.addItem(str(stud[1]))

        except con.Error as e:

            print("Select error" + e)


    def save_attendance(self):

        try:

            mydb = con.connect(host="localhost",port = 3307, user="root", password="", db="school")

            cursor = mydb.cursor()

            registration_number = self.cb41.currentText()

            attendance_date = self.tb41.text()

            start = self.tb42.text()

            duration = self.tb43.text()


            qry = "INSERT INTO attendance (registration_number,attendance_date,start,duration) VALUES(%s,%s,%s,%s)"

            value = (registration_number, attendance_date, start, duration)

            cursor.execute(qry, value)

            mydb.commit()

            self.tb41.setText("")

            self.tb42.setText("")

            self.tb43.setText("")

            self.l41.setText("Attendance Details Saved successfully!")

            QMessageBox.information(self, "School Management System", "Attendance Details Saved successfully!")

            cursor.execute(f"select full_name,email from student where registration_number = {registration_number}")
            res = cursor.fetchall()
            res = res[0]
            send_mail(res[0],attendance_date,start,duration,res[1])

            self.tabWidget.setCurrentIndex(1)

        except con.Error as e:
            print(e)
            self.l41.setText("Error in attendance details save !")


    def fill_date_on_combobox_selected(self):

        try:

            self.cb43.clear()

            mydb = con.connect(host="localhost",port = 3307, user="root", password="", db="school")

            cursor = mydb.cursor()

            cursor.execute("SELECT * FROM attendance where registration_number = '" + self.cb42.currentText() + "'")

            result = cursor.fetchall()

            print(result)

            if result:

                for exam in result:

                    self.cb43.addItem(str(exam[2]))

        except con.Error as e:

            print("Select error" + e)


    def fill_attendance_on_button_getattendance_clicked(self):

        try:

            mydb = con.connect(host="localhost",port = 3307, user="root", password="", db="school")

            cursor = mydb.cursor()

            cursor.execute("SELECT * FROM attendance where registration_number = '" + self.cb42.currentText() + "' and attendance_date = '" + self.cb43.currentText() + "'")

            result = cursor.fetchall()

            if result:

                for att in result:

                    self.tb44.setText(str(att[3]))

                    self.tb45.setText(str(att[4]))

        except con.Error as e:

            print("Select error" + e)


    def update_attendance(self):

        try:

            mydb = con.connect(host="localhost",port = 3307, user="root", password="", db="school")

            cursor = mydb.cursor()

            registration_number = self.cb42.currentText()

            attendance_date = self.cb43.currentText()

            morning = self.tb44.text()

            evening = self.tb45.text()

            qry = "UPDATE attendance set morning = '" + morning + "',evening= '" + evening + "' where registration_number ='" + registration_number + "' and attendance_date ='"+ attendance_date +"' "

            print(qry)

            cursor.execute(qry)

            self.tb44.setText("")

            self.tb45.setText("")


            mydb.commit()

            self.l42.setText("Attendance Details Modified successfully!")

            QMessageBox.information(self, "School Management System", "Attendance Details Modified successfully!")

            self.tabWidget.setCurrentIndex(1)

        except con.Error as e:

            self.l42.setText("Error in attendance modification!")


    def delete_attendance(self):

        m=QMessageBox.question(self,"Delete","Are you sure want to delete",QMessageBox.Yes|QMessageBox.No)

        if m == QMessageBox.Yes:

            try:

                mydb = con.connect(host="localhost",port = 3307, user="root", password="", db="school")

                cursor = mydb.cursor()

                registration_number = self.cb42.currentText()

                attendance_date = self.cb43.currentText()

                qry = "DELETE from attendance where registration_number ='" + registration_number + "' and attendance_date = '" + attendance_date + "' "

                cursor.execute(qry)

                mydb.commit()

                self.tb44.setText("")

                self.tb45.setText("")

                self.l42.setText("Attendance Details Deleted successfully!")

                QMessageBox.information(self, "School Management System", "Attendance Details Deleted successfully!")

                self.tabWidget.setCurrentIndex(1)

            except con.Error as e:

                self.l42.setText("Error in attendance deletion!")


    ######## FEES FORM ##########


    def show_fees(self):

        self.tabWidget.setCurrentIndex(6)

        self.fill_student_regno_combobox_fees()

        self.fill_receipt_number()

        self.fill_next_receipt_number()

        self.tb54.setText(str(date.today()))


    def fill_next_receipt_number(self):

        try:

            rn=0

            mydb = con.connect(host="localhost",port = 3307, user="root", password="", db="school")

            cursor = mydb.cursor()

            cursor.execute("SELECT * FROM fees")

            result = cursor.fetchall()

            if result:

                for stud in result:

                    rn += 1

            self.tb51.setText(str(rn+1))

        except con.Error as e:
            print(e)
            print("Select error" + e)



    def fill_student_regno_combobox_fees(self):

        try:

            mydb = con.connect(host="localhost",port = 3307, user="root", password="", db="school")

            cursor = mydb.cursor()

            cursor.execute("SELECT * FROM student")

            self.cb51.clear()

            result = cursor.fetchall()

            if result:

                for stud in result:

                    self.cb51.addItem(str(stud[1]))

        except con.Error as e:

            print("Select error" + e)


    def fill_receipt_number(self):

        try:

            mydb = con.connect(host="localhost",port = 3307, user="root", password="", db="school")

            cursor = mydb.cursor()

            cursor.execute("SELECT * FROM fees")

            self.cb52.clear()

            result = cursor.fetchall()

            if result:

                for stud in result:

                    self.cb52.addItem(str(stud[1]))

        except con.Error as e:

            print("Select error" + e)


    def save_fees(self):

        try:

            mydb = con.connect(host="localhost",port = 3307, user="root", password="", db="school")

            cursor = mydb.cursor()

            receipt_number = self.tb51.text()

            registration_number = self.cb51.currentText()
        
            reason = self.tb52.text()

            amont = self.tb53.text()

            fees_date = self.tb54.text()

            value = (receipt_number,registration_number,reason,amont,fees_date)
            


            cursor.execute("""INSERT INTO fees (receipt_number,registration_number,reason,amont,fees_date) VALUES(%s,%s,%s,%s,%s)""",(value))
            mydb.commit()


            self.l81.setText(self.tb51.text())

            self.l82.setText(self.tb54.text())

            cursor.execute("SELECT * FROM student where registration_number='"+ registration_number +"'")

            result = cursor.fetchone()

            if result:

                print(result)

            self.l83.setText(str(result[2]))

            self.l84.setText(self.tb53.text())

            self.l85.setText(self.tb52.text())

            self.l86.setText(self.tb54.text())


            self.tb51.setText("")

            self.tb52.setText("")

            self.tb53.setText("")

            self.tb54.setText("")

            self.fill_receipt_number()


            self.l51.setText("Fees Details Saved successfully!")

            QMessageBox.information(self, "School Management System", "Fees Details Saved successfully!")

            self.tabWidget.setCurrentIndex(8)

        except con.Error as e:
            print(e)
            self.l51.setText("Error in fees details save !")


    def fill_fees_details_on_receipt_number_selected(self):

        try:

            mydb = con.connect(host="localhost",port = 3307, user="root", password="", db="school")

            cursor = mydb.cursor()
            cursor.execute("SELECT * FROM fees where receipt_number = '" + self.cb52.currentText() + "' ")
            result = cursor.fetchall()
            print(result)
            if result:

                for att in result:

                    self.tb55.setText(str(att[2]))

                    self.tb56.setText(str(att[3]))

                    self.tb57.setText(str(att[4]))

                    self.tb58.setText(str(att[5]))

        except con.Error as e:
            print(e)
            print("Select error" + e)


    def update_fees(self):

        try:

            mydb = con.connect(host="localhost",port = 3307, user="root", password="", db="school")

            cursor = mydb.cursor()

            receipt_number = self.cb52.currentText()

            registration_number = self.tb55.text()

            reason = self.tb56.text()

            amont = self.tb57.text()

            fees_date = self.tb58.text()

            qry = "UPDATE fees set registration_number = '" + registration_number + "',reason= '" + reason + "' ,amont= '" + amont + "' ,fees_date= '" + fees_date + "' where receipt_number ='" + receipt_number + "' "

            print(qry)

            cursor.execute(qry)

            mydb.commit()


            self.l81.setText(receipt_number)

            self.l82.setText(fees_date)

            cursor.execute("SELECT * FROM student where registration_number='"+ registration_number +"'")

            result = cursor.fetchone()

            if result:

                print(result)

            self.l83.setText(str(result[2]))

            self.l84.setText(amont)

            self.l85.setText(reason)

            self.l86.setText(fees_date)



            self.tb55.setText("")

            self.tb56.setText("")

            self.tb57.setText("")

            self.tb58.setText("")


            self.l52.setText("Fees Details Modified successfully!")

            QMessageBox.information(self, "School Management System", "Fees Details Updated successfully!")

            self.tabWidget.setCurrentIndex(8)

        except con.Error as e:
            print(e)
            self.l52.setText("Error in fees modification!")


    def delete_fees(self):

        m=QMessageBox.question(self,"Delete","Are you sure want to delete",QMessageBox.Yes|QMessageBox.No)

        if m == QMessageBox.Yes:

            try:

                mydb = con.connect(host="localhost",port = 3307, user="root", password="", db="school")

                cursor = mydb.cursor()

                receipt_number = self.cb52.currentText()

                qry = "DELETE from fees where receipt_number ='" + receipt_number + "' "

                cursor.execute(qry)

                mydb.commit()

                self.tb55.setText("")

                self.tb56.setText("")

                self.tb57.setText("")

                self.tb58.setText("")

                self.fill_receipt_number()

                self.l42.setText("Fees Details Deleted successfully!")

                QMessageBox.information(self, "School Management System", "Fees Details Deleted successfully!")

                self.tabWidget.setCurrentIndex(1)

            except con.Error as e:

                self.l42.setText("Error in Fees deletion!")


    ######## REPORT FORM ##########


    def show_report(self):

        table_name = self.sender()

        print(table_name.text())

        self.l61.setText(table_name.text())

        self.tabWidget.setCurrentIndex(7)

        try:

            self.tableReport.setRowCount(0)

            if(table_name.text()=="Student Reports"):


                mydb = con.connect(host="localhost",port = 3307, user="root", password="", db="school")

                cursor = mydb.cursor()

                cursor.execute("SELECT registration_number,full_name,gender,date_of_birth,age,adress,phone,email,standard FROM student")

                result = cursor.fetchall()

                r = 0

                c = 0

                for row_number, row_data in enumerate(result):

                    r += 1

                    c = 0

                    for column_number, data in enumerate(row_data):

                        c += 1

                self.tableReport.clear()

                self.tableReport.setColumnCount(c)

                for row_number, row_data in enumerate(result):

                    self.tableReport.insertRow(row_number)

                    for column_number, data in enumerate(row_data):

                        self.tableReport.setItem(row_number, column_number, QTableWidgetItem(str(data)))

                self.tableReport.setHorizontalHeaderLabels(['Register Number','Name','Gender','Date of birth','Age','Adress','Phone','Email','Standard'])


            if(table_name.text()=="Mark Reports"):

                mydb = con.connect(host="localhost",port = 3307, user="root", password="", db="school")

                cursor = mydb.cursor()

                cursor.execute("SELECT registration_number,exam_name,grade FROM mark")

                result = cursor.fetchall()

                r = 0

                c = 0

                for row_number, row_data in enumerate(result):

                    r += 1

                    c = 0

                    for column_number, data in enumerate(row_data):

                        c += 1

                self.tableReport.clear()

                self.tableReport.setColumnCount(c)

                for row_number, row_data in enumerate(result):

                    self.tableReport.insertRow(row_number)

                    for column_number, data in enumerate(row_data):

                        self.tableReport.setItem(row_number, column_number, QTableWidgetItem(str(data)))

                self.tableReport.setHorizontalHeaderLabels(['Register Number', 'Exam Name', 'Language', 'English', 'Maths', 'Science', 'Social'])


            if(table_name.text()=="Attendance Reports"):

                mydb = con.connect(host="localhost",port = 3307, user="root", password="", db="school")

                cursor = mydb.cursor()

                cursor.execute("SELECT registration_number,attendance_date,start,duration FROM attendance")

                result = cursor.fetchall()

                r = 0

                c = 0

                for row_number, row_data in enumerate(result):

                    r += 1

                    c = 0

                    for column_number, data in enumerate(row_data):

                        c += 1

                self.tableReport.clear()

                self.tableReport.setColumnCount(c)

                for row_number, row_data in enumerate(result):

                    self.tableReport.insertRow(row_number)

                    for column_number, data in enumerate(row_data):

                        self.tableReport.setItem(row_number, column_number, QTableWidgetItem(str(data)))

                self.tableReport.setHorizontalHeaderLabels(['Register Number', 'Attendance Date', 'Morning', 'Evening'])


            if(table_name.text()=="Fees Reports"):

                mydb = con.connect(host="localhost",port = 3307, user="root", password="", db="school")

                cursor = mydb.cursor()

                cursor.execute("SELECT receipt_number,registration_number,reason,amont,fees_date FROM fees")

                result = cursor.fetchall()

                r = 0

                c = 0

                for row_number, row_data in enumerate(result):

                    r += 1

                    c = 0

                    for column_number, data in enumerate(row_data):

                        c += 1

                self.tableReport.clear()

                self.tableReport.setColumnCount(c)

                for row_number, row_data in enumerate(result):

                    self.tableReport.insertRow(row_number)

                    for column_number, data in enumerate(row_data):

                        self.tableReport.setItem(row_number, column_number, QTableWidgetItem(str(data)))

                self.tableReport.setHorizontalHeaderLabels(['Receipt Number', 'Register Number', 'Reason', 'amont', 'Paid Date'])


        except con.Error as e:
            print(e)
            print("Select error")


    ###### Export to Excel ##############

   

    def export_table(self):

        print("Export started")

        columnHeaders = []

        for j in range(self.tableReport.model().columnCount()):

            columnHeaders.append(self.tableReport.horizontalHeaderItem(j).text())

            print(columnHeaders)


        df = pd.DataFrame(columns = columnHeaders)


        for row in range(self.tableReport.rowCount()):

            for col in range(self.tableReport.columnCount()):

               df.at[row,columnHeaders[col]] = self.tableReport.item(row,col).text()


        xlname = self.l61.text() + str(date.today()) + ".xlsx"

       

        df.to_excel(xlname,index=False)


        QMessageBox.information(self, "School Management System", "File Exported successfully! " + xlname)



    ###### PRINT ########################

    def print_file(self):

        printer = QPrinter(QPrinter.HighResolution)

        dialog = QPrintDialog(printer,self)

        if dialog.exec_() == QPrintDialog.Accepted:

            self.tb1.print_(printer)


    def cancel_print(self):

        self.tabWidget.setCurrentIndex(1)


    ############# Logout ##########

   

    def logout(self):

        self.menubar.setVisible(False)

        self.tb01.setText("")

        self.tb02.setText("")         

        self.tabWidget.setCurrentIndex(0)

        QMessageBox.information(self, "School management system", "You are loged out successfully")

       



def main():

    app = QApplication(sys.argv)

    window = MainApp()

    window.show()

    app.exec_()



if __name__ == '__main__':

    main()



