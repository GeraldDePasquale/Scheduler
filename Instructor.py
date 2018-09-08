import time
from datetime import datetime, time
from Importer import Importer

directory = "C:\\ProgramData\\MathnasiumScheduler"
FILEOPENOPTIONS = dict(defaultextension='.csv', filetypes=[('XLSX', '*.xlsx'), ('CSV file', '*.csv')])

# define days consistent with datetime.weekday() for printing and sorting purposes
# ToDo refactor common data
mon = 0
tue = 1
wed = 2
thu = 3
fri = 4
sat = 5
sun = 6

mon2thu = [mon, tue, wed, thu]

instructors = []
initialized = False
unavailable = time(0, 0)

# center instruction hours
# ToDo refactor common data
instructionHours = {sun: [], \
                    mon: [time(15, 30), time(19, 30)], \
                    tue: [time(15, 30), time(19, 30)], \
                    wed: [time(15, 30), time(19, 30)], \
                    thu: [time(15, 30), time(19, 30)], \
                    fri: [], \
                    sat: [time(10, 00), time(14, 00)]}

slot01 = [time(13, 00), time(13, 30)]
slot02 = [time(13, 30), time(14, 00)]
slot03 = [time(14, 00), time(14, 30)]
slot04 = [time(14, 30), time(15, 00)]
slot05 = [time(15, 00), time(15, 30)]
slot06 = [time(15, 30), time(16, 00)]
slot07 = [time(16, 00), time(16, 30)]
slot08 = [time(16, 30), time(17, 00)]
slot09 = [time(17, 00), time(17, 30)]
slot10 = [time(17, 30), time(18, 00)]
slot11 = [time(18, 00), time(18, 30)]
slot12 = [time(18, 30), time(19, 00)]
slot13 = [time(19, 00), time(19, 30)]
slot14 = [time(19, 30), time(20, 00)]

weekdayslots = [slot01, slot02, slot03, slot04, slot05, slot06, slot07, \
                slot08, slot09, slot10, slot11, slot12, slot13, slot14]

slot15 = [time(9, 00), time(9, 30)]
slot16 = [time(9, 30), time(10, 00)]
slot17 = [time(10, 00), time(10, 30)]
slot18 = [time(10, 30), time(11, 00)]
slot19 = [time(11, 00), time(11, 30)]
slot20 = [time(11, 30), time(12, 00)]
slot21 = [time(12, 00), time(12, 30)]
slot22 = [time(12, 30), time(13, 00)]
slot23 = [time(13, 00), time(13, 30)]
slot24 = [time(13, 30), time(14, 00)]
slot25 = [time(14, 00), time(14, 30)]
slot26 = [time(14, 30), time(15, 00)]

weekendslots = [slot15, slot16, slot17, slot18, slot19, slot20, slot21, \
                slot22, slot23, slot24, slot25, slot26]

# number of virtual instructors (move to config file)
virtual_instructors = 15

# center instruction hours
zero_hours = [time(0, 0), time(0, 0)]
twenty_four_hours = [time(0, 0), time(23, 59)]
# virtual_instructor_availability = \
#                     {sun: twenty_four_hours, \
#                      mon: twenty_four_hours, \
#                      tue: twenty_four_hours, \
#                      wed: twenty_four_hours, \
#                      thu: twenty_four_hours, \
#                      fri: twenty_four_hours, \
#                      sat: twenty_four_hours}

virtual_instructor_availability = \
    {sun: zero_hours, \
     mon: zero_hours, \
     tue: zero_hours, \
     wed: zero_hours, \
     thu: zero_hours, \
     fri: zero_hours, \
     sat: zero_hours}


class Instructor:
    name = None
    schedule = None
    email_address = None
    availability = None
    availability_reported = None
    # ToDo set calendar_id to employee's google calendar id
    calendar_id = None
    rank = 1000
    cost = 15
    max_hours_per_week = 20
    min_hours_per_day = 2
    work_day_preferences = {sun: 0, mon: 2, tue: 3, wed: 4, thu: 5, fri: 6, sat: 7}
    dayStrings = {sun: "Sunday", mon: "Monday", tue: "Tuesday", \
                  wed: "Wednesday", thu: "Thursday", fri: "Friday", \
                  sat: "Saturday"}
    isScheduled = False  # ToDo consider renaming isActive
    isVirtual = False
    # Alternative Scheduling Option
    minActivationEvent = None  # min start time
    maxActivationEvent = None  # max stop time from availaility
    activationEvent = None  # start time
    deactivationEvent = None

    def __init__(self):
        Instructor.rank += 1

    @staticmethod
    def create_real_instructor_from_row(log_file, row=None):
        real_instructor = Instructor()

        # define work schedule
        real_instructor.schedule = {sun: [], mon: [], tue: [], wed: [], thu: [], fri: [], sat: []}

        # set instructor name
        real_instructor.availability_reported = row[0].value
        real_instructor.email_address = row[1].value
        real_instructor.name = row[2].value

        # set instructors hours of availability based upon the lines in the InstructorAvailability input file
        real_instructor.availability = {sun: (unavailable, unavailable), \
                                        mon: (row[6].value, row[7].value), \
                                        tue: (row[9].value, row[10].value), \
                                        wed: (row[12].value, row[13].value), \
                                        thu: (row[15].value, row[16].value), \
                                        fri: (unavailable, unavailable), \
                                        sat: (row[18].value, row[19].value)}

        # set rank
        real_instructor.rank = 999  # rank not found in data
        # find row of the instructor, then assign rank from cell
        # cross references email address in availability file to email address in configuration sheet then assigns rank
        first_row = 2  # skip the headers
        last_row = Importer.config_ws.max_row
        last_col = Importer.config_ws.max_column
        for row in Importer.config_ws.iter_rows(min_row=first_row, max_col=last_col, max_row=last_row):
            if real_instructor.email_address == row[2].value:
                real_instructor.rank = row[3].value

        # set cost
        real_instructor.cost = 15

        # set max hours per week
        real_instructor.maxHrsPerWeek = 20

        # set min hours per day
        real_instructor.minHrsPerDay = 2

        # set work day preferences
        real_instructor.workDayPreferences = {sun: 0, mon: 2, tue: 3, wed: 4, thu: 5, fri: 6, sat: 7}

        # set dayofweek strings
        real_instructor.dayStrings = {sun: "Sunday", mon: "Monday", tue: "Tuesday", \
                                      wed: "Wednesday", thu: "Thursday", fri: "Friday", \
                                      sat: "Saturday"}

        # set isSchedule flag - this variable is set by the scheduling process
        real_instructor.isScheduled = False

        real_instructor.isVirtual = False

        return real_instructor

    @staticmethod
    def create_virtual_instructor(index=1, day=None, workSlot=None):
        virtual_instructor = Instructor()
        virtual_instructor.schedule = {sun: [], mon: [], tue: [], wed: [], thu: [], fri: [], sat: []}
        virtual_instructor.availability_reported = datetime.now()
        # ToDo remove hard coded address
        virtual_instructor.email_address = "stafford@mathnasium.com"
        virtual_instructor.name = 'Gap ' + virtual_instructor.dayString(day) + " " + workSlot[0].strftime("%H:%M") + \
                                  " to " + workSlot[1].strftime("%H:%M")
        virtual_instructor.availability = {sun: zero_hours, mon: zero_hours, tue: zero_hours, wed: zero_hours, \
                                           thu: zero_hours, fri: zero_hours, sat: zero_hours}
        virtual_instructor.availability[day] = workSlot
        virtual_instructor.rank = Instructor.rank
        virtual_instructor.cost = 15
        virtual_instructor.maxHrsPerWeek = 100
        virtual_instructor.minHrsPerDay = 0
        virtual_instructor.workDayPreferences = {sun: 0, mon: 2, tue: 3, wed: 4, thu: 5, fri: 6, sat: 7}
        virtual_instructor.dayStrings = {sun: "Sunday", mon: "Monday", tue: "Tuesday", \
                                         wed: "Wednesday", thu: "Thursday", fri: "Friday", \
                                         sat: "Saturday"}
        virtual_instructor.isScheduled = False
        virtual_instructor.isVirtual = True
        return virtual_instructor

    @staticmethod
    def create_instructors(run_log):

        # Update Availability Work Sheet: When instructor not available, change none to time(0,0)
        first_row = 2  # skip the headers
        last_row = Importer.instructor_availability_ws.max_row
        last_col = Importer.instructor_availability_ws.max_column
        for row in Importer.instructor_availability_ws.iter_rows(min_row=first_row, max_col=last_col, max_row=last_row):
#            for i in [5, 6, 8, 9, 11, 12, 14, 15, 17, 18]:
            for i in [6, 7, 9, 10, 12, 13, 15, 16, 18, 19]:
                if row[i].value == None: row[i].value = time(0, 0)
        # for eachInstructor in instructors: print(eachInstructor.name)

        # Create Instructors
        Instructor.instructors = []
        print("\nCreating Instructors (from Instructor Availability and Configuration File")
        for row in Importer.instructor_availability_ws.iter_rows(min_row=first_row, max_col=last_col, max_row=last_row):
            Instructor.instructors.append(Instructor.create_real_instructor_from_row(run_log, row))
        # Add Virtual Instructors to identify periods that cannot covered by existing staff
        # for i in range (1, virtual_instructors):
        #     Instructor.instructors.append(Instructor.create_virtual_instructor(i))
        i = 0
        maxSlotsNeeded = 5  # ToDo hard coded
        for j in range(1, maxSlotsNeeded):
            for day in mon2thu:
                for slot in weekdayslots:
                    i = i + 1
                    Instructor.instructors.append(Instructor.create_virtual_instructor(i, day, slot))
            for slot in weekendslots:
                i = i + 1
                Instructor.instructors.append(Instructor.create_virtual_instructor(i, sat, slot))

        # for eachInstructor in Instructor.instructors: print("\tInstructor: ", eachInstructor.name,
        #                                                              "Rank: ", eachInstructor.rank)
        return Instructor.instructors

    # define instructor sort
    def __lt__(self, other):
        return self.rank < other.rank

    def __str__(self):
        return str(self.name)

    def __print__(self):
        print(str(self.name, self.email_address, self.rank))

    # weekdayString(self, dayofweek) returns string representation of the day
    def dayString(self, integer):
        return self.dayStrings[integer]

    # isAvailable returns true if the instructor is available to work at the time of the event
    def isAvailable(self, event):
        dayNeeded = event.event_time.weekday()
        timeNeeded = event.event_time.time()
        imScheduled = self.isScheduled
        availableStartTime = self.availability[dayNeeded][0]
        availableStopTime = self.availability[dayNeeded][1]
        result = (not imScheduled) and \
                 (timeNeeded >= availableStartTime) and \
                 (timeNeeded < availableStopTime)
        return result

    # isAvailableToOpen returns true if the instructor is available to start work when the center opens
    def isAvailableToOpen(self, event):
        dayNeeded = event.event_time.weekday()
        timeNeeded = instructionHours[dayNeeded][0]
        imScheduled = self.isScheduled
        myStartTime = self.availability[dayNeeded][0]
        myStopTime = self.availability[dayNeeded][1]
        result = (not imScheduled) and \
                 (timeNeeded >= myStartTime) and \
                 (timeNeeded < myStopTime)
        return result

    def isVirtual(self):
        return self.isVirtual

    # startWork adds the start work event time to the instructor's work schedule
    def startWork(self, event):
        dayNeeded = event.event_time.weekday()
        timeNeeded = event.event_time.time()
        self.schedule[dayNeeded].append(timeNeeded)
        self.isScheduled = True

    # startWorkWhenOpen adds start at open time to the instructor's work schedule
    def startWorkWhenOpen(self, event):
        dayNeeded = event.event_time.weekday()
        timeNeeded = instructionHours[dayNeeded][0]
        self.schedule[dayNeeded].append(timeNeeded)
        self.isScheduled = True

    # stopWork adds the stop work event time to the instructor's work schedule
    def stopWork(self, event):
        self.schedule[event.event_time.weekday()].append(event.event_time.time())
        self.isScheduled = False

    # mustDepart returns true if the event time exceeds the instructors availability
    def mustDepart(self, event):
        dayNeeded = event.event_time.weekday()
        timeNeeded = event.event_time.time()
        myStopTime = self.availability[dayNeeded][1]
        result = (timeNeeded >= myStopTime)
        return result

    # departWork adds stop work event but leaves instructor in isScheduled state
    def departWork(self, event):
        self.schedule[event.event_time.weekday()].append(event.event_time.time())
        self.isScheduled = True

    # hoursScheduled returns the number of hours the employee is scheduled to work
    #    def hoursScheduledNow(self,event):

    # adjust time to land on the hour or half hour
    def adjustTime(self, aTime):
        if aTime.minute < 30:
            return time(aTime.hour, 0)
        elif aTime.minute > 30:
            return time(aTime.hour, 30)
        return aTime

    # finalizeSchedule modifies the schedule so that each work day includes one start and one stop time
    def finalizeSchedule(self):
        for eachKey in self.schedule.keys():
            if self.schedule[eachKey]:
                myStartTime = max(self.schedule[eachKey][0], self.availability[eachKey][0])
                myStopTime = min(self.availability[eachKey][1], self.schedule[eachKey][len(self.schedule[eachKey]) - 1])
                # ToDo figure out why stop work is not set and fix if appropriate
                # Bug fix: If no stop work ordered, stop work at end of day. Better fix check for no stop work
                if myStopTime == myStartTime: myStopTime = instructionHours[eachKey][1]
                self.schedule[eachKey] = [self.adjustTime(myStartTime), self.adjustTime(myStopTime)]
        return self.schedule

    def work_day_tuple(self, day):
        return [self.name, self.dayString(day), self.schedule[day][0], self.schedule[day][1]]

    # create the an html message containing the instructor's schedule
    def schedule_as_html_msg(self):
        cr = '<br/>'
        sp = ' '
        fr = ' from '
        to = ' to '
        # ToDo refactor: salutation and message to editable configuration file
        salutation = 'Feel free to call me at 540-907-9306 if you have any questions.' + cr + cr + 'Jerry'
        msg = 'Dear' + sp + self.name + ',' + cr + cr + 'Here is your schedule:' + cr + cr
        for eachKey in self.schedule.keys():
            if self.schedule[eachKey]:
                today = self.dayString(eachKey) + fr + str(self.schedule[eachKey][0]) + to + str(
                    self.schedule[eachKey][1]) + cr
                msg = msg + today
        msg = msg + cr
        msg = msg + salutation + cr
        return msg

    # create the a text message containing the instructor's schedule
    def schedule_as_plain_msg(self):
        cr = '\n'
        sp = ' '
        fr = ' from '
        to = ' to '
        # ToDo refactor - salutation and msg to configuration file
        salutation = 'Feel free to call me at 540-907-9306 if you have any questions.' + cr + cr + 'Jerry'
        msg = 'Dear' + sp + self.name + ',' + cr + cr + 'Here is your schedule:' + cr + cr
        for eachKey in self.schedule.keys():
            if self.schedule[eachKey]:
                today = self.dayString(eachKey) + fr + str(self.schedule[eachKey][0]) + to + str(
                    self.schedule[eachKey][1]) + cr
                msg = msg + today
        msg = msg + cr
        msg = msg + salutation + cr
        return msg
