import xlrd,os
from datetime import datetime
import pandas as pd
#匯入模組(Module)
import sys
import xlwt

sys.setrecursionlimit(1000000)

path = "attendance.xlsx"

# determine if application is a script file or frozen exe
if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
elif __file__:
    application_path = os.path.dirname(__file__)

abs_path = os.path.join(application_path, path)
print(abs_path)

# import xlsx file
if not os.path.exists(abs_path):
    open(abs_path,str(abs_path))
else:
    attendance=xlrd.open_workbook(abs_path)
    # another ver.
    read_file=pd.read_excel(abs_path)
    m_table= pd.DataFrame(read_file)

    print(m_table)
    table= attendance.sheets()[0]



'''
col value 0 : 差勤號碼
col value 1 : 姓名
col value 2 : 日期
col value 3 : 上班時間
col value 4 : 下班時間
col value 5 : 簽到時間
col value 6 : 簽退時間
col value 7 : 遲到時間
col value 8 : 早退時間
col value 9 : 是否礦工
col value 10 : 實際工作時間（最多就是上班時間到下班時間）
col value 11 : 例外情況
col value 12 : 出勤時間 (簽到時間-簽退時間)
'''

'''
出勤表類別
'''
class Attendance:
    def __init__(self, index=0, name=0, date=0, workhour=0, offTime=0,
                 checkInTime=0, checkOutTime=0, lateTime=0, earlyDepartureTime=0,
                 isAbsenteeism=0, actualWorkTime=0, exception=0, worktime=0):
        self.index = index
        self.name = name
        self.date = date
        self.workhour = workhour
        self.offTime = offTime
        self.checkInTime = checkInTime
        self.checkOutTime = checkOutTime
        self.lateTime = lateTime
        self.earlyDepartureTime = earlyDepartureTime
        self.isAbsenteeism = isAbsenteeism
        self.actualWorkTime = actualWorkTime
        self.exception = exception
        self.worktime = worktime

    def showIndex(self):
        print(self.index)


'''
員工類別
'''
class Staff:
    def __init__(self, index=0, name=0, worktime=0,
                 late_time=0, early_departure_time=0,
                 weekday_overtime=0, weekend_overtime=0,
                 late_day_num=0, early_departure_day_num=0,
                 weekday_overtime_num=0,weekend_overtime_num=0,not_coming_day=0,
                 staff_attendance=[]):
        self.index = index
        self.name = name
        self.worktime = worktime
        self.late_time = late_time
        self.early_departure_time = early_departure_time
        self.weekday_overtime = weekday_overtime
        self.weekend_overtime = weekend_overtime
        self.late_day_num = late_day_num
        self.weekday_overtime_num = weekday_overtime_num
        self.weekend_overtime_num = weekend_overtime_num
        self.early_departure_day_num = early_departure_day_num
        self.staff_attendance = staff_attendance
        self.not_coming_day = not_coming_day
    
    def addNewAttendance(self, _attendance):
        self.staff_attendance.append(_attendance)

    def addLateNum(self):
        self.late_day_num += 1
        print(self.late_day_num)

    '''
    debug用
    '''
    def showWorktime(self):
        print(self.worktime)
        
    def print_offtime(self):
        for _staff_attendance in self.staff_attendance:
            print("offtime",self.name,datetime.strptime(_staff_attendance.offTime,"%H:%M"))
    
    def print_checkOutTime(self):
        for _staff_attendance in self.staff_attendance:
            if _staff_attendance.checkOutTime ==  '':
                _staff_attendance.checkOutTime=_staff_attendance.offTime
            print("checkOutTime",self.name,datetime.strptime(_staff_attendance.checkOutTime,"%H:%M"))
    
    '''
    計算遲到、早退、加班、工作時間以及其次數
    '''
    def calculate_all(self):
        self.calculate_worktime()
        self.calculate_late_time()
        self.calculate_early_departure_time()
        self.calculate_weekday_overtime()
        self.calculate_weekend_overtime()
        
    def calculate_worktime(self):
        self.worktime=0
        workday = 0
        self.not_coming_day=0
        for _staff_attendance in self.staff_attendance:
            
            if _staff_attendance.worktime == '':
                self.not_coming_day+=1
                _staff_attendance.worktime="0:0"
            if _staff_attendance.worktime != '' and _staff_attendance.worktime != "0:0":
                work_time=datetime.strptime(_staff_attendance.worktime,"%H:%M")
                hour=work_time.hour*60
                minute=work_time.minute
                self.worktime+=(hour+minute)
                workday+=1
        minute= self.worktime%60
        hour=self.worktime//60
        self.worktime=str(hour)+":"+str(minute)
        print("工作時間",self.name,self.worktime,"工作天數",workday,"請假/曠工/出差/還沒開始上班",self.not_coming_day)
    
    
    def calculate_late_time(self):
        self.late_time = 0
        late_num = 0
        for _staff_attendance in self.staff_attendance:
            # 過濾沒來的日子 and 假日
            work_date=datetime.strptime(_staff_attendance.date,"%Y/%m/%d")
            if _staff_attendance.worktime == '' or _staff_attendance.worktime == "0:0" and work_date.weekday() <= 6:
                continue
            
            if _staff_attendance.checkInTime ==  '':
                _staff_attendance.checkInTime=_staff_attendance.workhour
                
            #change time date to datetime type     
            workhour = datetime.strptime(_staff_attendance.workhour,"%H:%M")
            checkInTime = datetime.strptime(_staff_attendance.checkInTime,"%H:%M")
            offTime = datetime.strptime(_staff_attendance.offTime,"%H:%M")
            isLate = checkInTime.hour > 6 and (checkInTime.hour >= workhour.hour) and (checkInTime.hour < 13)
            if isLate:
                latetime = (checkInTime-workhour).seconds//60
                self.late_time+=(latetime)
                late_num+=1
        minute= self.late_time%60
        hour=self.late_time//60
        self.late_time=str(hour)+":"+str(minute)
        self.late_day_num = late_num
        print("平日遲到",self.name,self.late_time,"次數",self.late_day_num)
    

    
    def calculate_early_departure_time(self):
        self.early_departure_time = 0
        early_departure_time_num = 0
        for _staff_attendance in self.staff_attendance:
            # 過濾沒來的日子 and 假日
            work_date=datetime.strptime(_staff_attendance.date,"%Y/%m/%d")
            if _staff_attendance.worktime == '' or _staff_attendance.worktime == "0:0" and work_date.weekday() <= 6:
                continue
            
            if _staff_attendance.checkOutTime ==  '':
                _staff_attendance.checkOutTime=_staff_attendance.offTime
                
            #change time date to datetime type     
            checkOutTime = datetime.strptime(_staff_attendance.checkOutTime,"%H:%M")
            offTime = datetime.strptime(_staff_attendance.offTime,"%H:%M")
            isEarlyDeparture = (checkOutTime.hour > 13) and (checkOutTime.hour < offTime.hour)
            if isEarlyDeparture:
                earlyDepartureTime = (offTime-checkOutTime).seconds//60
                self.early_departure_time+=(earlyDepartureTime)
                early_departure_time_num+=1
        minute= self.early_departure_time%60
        hour=self.early_departure_time//60
        self.early_departure_time=str(hour)+":"+str(minute)
        self.early_departure_day_num = early_departure_time_num
        print("平日早退",self.name,self.early_departure_time,"次數",self.early_departure_day_num)
                
        
    def calculate_weekday_overtime(self):
        self.weekday_overtime=0
        overtime_num = 0
        for _staff_attendance in self.staff_attendance:
            work_date=datetime.strptime(_staff_attendance.date,"%Y/%m/%d")
            if _staff_attendance.worktime == '' or _staff_attendance.worktime == "0:0" and work_date.weekday() <= 6:
                continue
            if _staff_attendance.checkOutTime ==  '':
                _staff_attendance.checkOutTime=_staff_attendance.offTime
            #change time date to datetime type     
            offtime = datetime.strptime(_staff_attendance.offTime,"%H:%M")
            checkouttime = datetime.strptime(_staff_attendance.checkOutTime,"%H:%M")
            isOvertime = checkouttime.hour < 6 or (checkouttime.hour >= offtime.hour)
            if isOvertime:
                overtime = (checkouttime-offtime).seconds//60
                self.weekday_overtime+=(overtime)
                overtime_num+=1
        minute= self.weekday_overtime%60
        hour=self.weekday_overtime//60
        self.weekday_overtime=str(hour)+":"+str(minute)
        self.weekday_overtime_num = overtime_num
        print("平日加班",self.name,self.weekday_overtime,"次數",self.weekday_overtime_num)
        
    def calculate_weekend_overtime(self):
        self.weekend_overtime = 0
        weekend_overtime_num = 0
        for _staff_attendance in self.staff_attendance:
            work_date=datetime.strptime(_staff_attendance.date,"%Y/%m/%d")
            if work_date.weekday() == 6 or work_date.weekday() == 7:
                if _staff_attendance.worktime == '':
                    _staff_attendance.worktime="0:0"
                if _staff_attendance.worktime != '':
                    work_time=datetime.strptime(_staff_attendance.worktime,"%H:%M")
                    hour=work_time.hour*60
                    minute=work_time.minute
                    self.weekend_overtime+=(hour+minute)
                    weekend_overtime_num+=1
        minute= self.weekend_overtime%60
        hour=self.weekend_overtime//60
        self.weekend_overtime=str(hour)+":"+str(minute)
        self.weekend_overtime_num=weekend_overtime_num
        print("假日加班",self.name,self.weekend_overtime,"次數",self.weekend_overtime_num)

def calculate_date_interval(date1,date2):
    date1 = datetime.strptime(date1,"%H:%M")
    date2 = datetime.strptime(date2,"%H:%M")
    minute = (d1-d2).seconds//60
    hour = minute//60
    minute %= 60 
    return str(hour) + ":" + str(minute)
    
      
# use first attendance to initial attribute in staff 
def create_new_Staff(first_attendance, staff_attedance):
    newStaff= Staff(
        first_attendance.index,
        first_attendance.name,
        first_attendance.worktime,
        late_time_, early_departure_time_, weekday_overtime_, weekend_overtime_, 
        late_day_num_, early_departure_day_num_,weekday_overtime_num_,weekend_overtime_num_,not_coming_day,
        staff_attedance)
    return newStaff

# create new staff attendance though table data
def create_new_staff_attendance(index):
    # 讀取該名員工第index筆出勤紀錄
    staff_attendanceList= table.row_values(index)
    newStaffAttendance= Attendance(
        staff_attendanceList[0], staff_attendanceList[1],
        staff_attendanceList[2], staff_attendanceList[3], staff_attendanceList[4],
        staff_attendanceList[5], staff_attendanceList[6], staff_attendanceList[7],
        staff_attendanceList[8], staff_attendanceList[9], staff_attendanceList[10],
        staff_attendanceList[11], staff_attendanceList[12])
    return newStaffAttendance

def calculate_all_staff_worktime():
    for staff in StaffList:
        staff.calculate_worktime()

def calculate_all_staff_weekend_overtime():
    for staff in StaffList:
        staff.calculate_weekend_overtime()
        
def calculate_all_staff_weekday_overtime():
    for staff in StaffList:
        staff.calculate_weekday_overtime()

def calculate_all_staff_late_time():
    for staff in StaffList:
        staff.calculate_late_time()

def calculate_all_staff_early_departure_time():
    for staff in StaffList:
        staff.calculate_early_departure_time()

def calculate_all_staff_all():
    for staff in StaffList:
        staff.calculate_all()

'''
建立excel
https://thai-lin.blogspot.com/2017/07/pythonexcel.html
'''
#建立Workbook物件
AttendanceAnalize= xlwt.Workbook(encoding="utf-8")
#使用Workbook裡的add_sheet函式來建立Worksheet
sheet1 = AttendanceAnalize.add_sheet("AttendanceAnalize")

def createFile(orig_args):
    filename = os.path.join(application_path,"AttendanceAnalize.xls")
    output(filename,StaffList)
'''
name=0, worktime=0,
late_time=0, early_departure_time=0,
weekday_overtime=0, weekend_overtime=0,
late_day_num=0, early_departure_day_num=0,
weekday_overtime_num=0,weekend_overtime_num=0,not_coming_day=0,
'''
def output(filename,_staffList):
    #使用Worksheet裡的write函式將值寫入
    sheet1.write(0,0,"姓名")
        
    #worktime & latetime & early departure time
    sheet1.write(0,1,"工作時間(小時：分鐘)")
    sheet1.write(0,2,"遲到時間(小時：分鐘)")
    sheet1.write(0,3,"早退時間(小時：分鐘)")
        
    #overtime (weekday and weekend)
    sheet1.write(0,4,"平日加班時間(小時：分鐘)")
    sheet1.write(0,5,"假日加班時間(小時：分鐘)")
        
    #worktime & latetime & early departure time num
    sheet1.write(0,6,"遲到次數")
    sheet1.write(0,7,"早退次數")
    sheet1.write(0,8,"平日加班次數")
        
    #overtime (weekday and weekend)
    sheet1.write(0,9,"假日加班次數")
    sheet1.write(0,10,"請假/曠工/出差/還沒開始上班")
    
    index  = 1
    for staff in _staffList:
        #name
        sheet1.write(index,0,staff.name)
        
        #worktime & latetime & early departure time
        sheet1.write(index,1,staff.worktime)
        sheet1.write(index,2,staff.late_time)
        sheet1.write(index,3,staff.early_departure_time)
        
        #overtime (weekday and weekend)
        sheet1.write(index,4,staff.weekday_overtime)
        sheet1.write(index,5,staff.weekend_overtime)
        
        #worktime & latetime & early departure time num
        sheet1.write(index,6,staff.late_day_num)
        sheet1.write(index,7,staff.early_departure_day_num)
        sheet1.write(index,8,staff.weekday_overtime_num)
        
        #overtime (weekday and weekend)
        sheet1.write(index,9,staff.weekend_overtime_num)
        sheet1.write(index,10,staff.not_coming_day)
        
        index += 1
    #將Workbook儲存為原生Excel格式的檔案
    AttendanceAnalize.save(filename)

def main():
    index = 0
    for name in table.col_values(1):
        if name != '姓名' and NameList.count(name) == 0:
            NameList.append(name)# 新增員工出勤紀錄表
            staff_attedance = []
            staff_attedance.append(create_new_staff_attendance(index))# 新增員工資料
            newStaff = create_new_Staff(staff_attedance[0], staff_attedance)
            StaffList.append(newStaff)
    
        elif NameList.count(name) > 0:
            #add new attendance of existence staff
            for staff in StaffList:
                if staff.name == name:
                    staff.addNewAttendance(create_new_staff_attendance(index))
        index+= 1
    calculate_all_staff_all()
    createFile(sys.argv)


# list
NameList = []
StaffList = []

# staff attribute
worktime_ = table.row_values(0)[3]
late_time_ = table.row_values(0)[3]
early_departure_time_ = table.row_values(0)[3]
weekday_overtime_ = table.row_values(0)[3]
weekend_overtime_ = table.row_values(0)[3]
late_day_num_ = 0
early_departure_day_num_ = 0
weekday_overtime_num_=0
weekend_overtime_num_=0
not_coming_day=0

# process main func
main()



