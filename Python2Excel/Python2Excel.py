#################################################################       
# THIS PROGRAM WILL KEEP TRACK OF HOURS WORKED                  # 
# MUST ENTER WEEKLY HOURS AND THEN COPY TO MONTHLY SPREADSHEET  #                                           
# Author: Shane Patrick Cronk                                   #
# Date:   6/14/2016                                             #
#                                                               #
#################################################################

import sys
import math
import xlsxwriter   #Excel Module
import tkinter   #GUI Module
import graphics  #Graphics Module


class TimeSheet():

    def __init__(self,parent):
        #keep reference to original "Parent"
        self.parent = parent

        #create a new Excel file and add a worksheet.
        self.Prog_Info()
        WeeklyTimesheet = input("""Please enter a filename for this weeks timesheet as 'Week#MonthYY.xlsx': """)
        self.workbook = xlsxwriter.Workbook(WeeklyTimesheet)
        self.worksheet1 = self.workbook.add_worksheet()
        #winden the first column to make the text clearer
        self.worksheet1.set_column('A:D',20)
        #Add a bold format to use the highlight
        self.bold = self.workbook.add_format({'bold':True})
        self.format = self.workbook.add_format()
        self.format.set_align('center')
        #self.format.set_color('blue')

    def write_times(self):
        #Set Row and Col To Zero So we have a starting location
        


        #Set COLUMN HEADINGS
        self.worksheet1.write('A1', 'NAME',self.bold)
        self.worksheet1.write('B1', 'DAY',self.bold)
        self.worksheet1.write('C1', 'DATE',self.bold)
        self.worksheet1.write('D1', 'HOURS WORKED',self.bold)
        self.worksheet1.write('E1', 'LOCATION',self.bold)
        #self.worksheet1.write('F1', 'TOTAL',self.bold)
        
        #Get user name and starting ROW number
        #ROW number should increment by 7's
        #This will take the user from Monday to Sunday, start again on Monday
        self.get_name()
        
        self.weekly = 7
        for i in range(self.weekly):
            self.getWeek_info()
        
        #self.total_it()

        #Close the workbook
        #try: NO LONGER REQUIRED BECAUSE ONLY USING A WEEKLY FORMAT
        self.workbook.close()
        #except:
            #pass

    def get_name(self):
        self.col = 0
        self.uName = input("Please enter your name: ")
        self.row_start = int(input("Starting Row?: "))
        self.worksheet1.write(self.row_start,self.col,self.uName)

        
    def getWeek_info(self):
        self.col = 0
        self.uDay = input("Please enter the day of the week: ")
        self.uDate = input("Please enter the date, MM/DD/YY: ")
        self.uHours = input ("Please enter the hours worked: ")
        self.uLocation = input("Please enter the location: ")
        

        #Write Values to worksheet
        self.worksheet1.write(self.row_start,self.col + 1,self.uDay)
        self.worksheet1.write(self.row_start,self.col + 2,self.uDate)
        self.worksheet1.write(self.row_start,self.col + 3,float(self.uHours))
        self.worksheet1.write(self.row_start,self.col + 4,self.uLocation)
        print("Thank You, Hours Recorded For: " + self.uDate + '\n')
        #Increment row_start by 1. 
        #This function will execute 7 times and store a week of data.
        self.row_start +=1

        #Total hours and write value to spreadsheet
    '''def total_it(self):
        self.worksheet1.write(self.row_start + 6,'E9','= D2 + D3 + D4 + D5 + D6')
    '''

        
    def Prog_Info(self):
        print("""Welcome to the timesheet. 
This program will save one weeks worth of hours in a spreadsheet.
The user must save the weekly spreadsheets into a monthly folder.
The user will begin with Monday as the first entry and continue until Sunday.""")
        print("""Author: Shane Cronk
Date: June,15,2016""" + '\n')

def main():


    newTimeSheet = TimeSheet(None)
    newTimeSheet.write_times()

main()