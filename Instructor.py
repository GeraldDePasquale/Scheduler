import time
from datetime import datetime, time
from Importer import Importer

directory = "C:\\ProgramData\\MathnasiumScheduler"
FILEOPENOPTIONS = dict(defaultextension='.csv', filetypes=[('XLSX', '*.xlsx'), ('CSV file', '*.csv')])

# define days consistent with datetime.weekday() for printing and sorting purposes
mon = 0
tue = 1
wed = 2
thu = 3
fri = 4
sat = 5
sun = 6

instructors = []
initialized = False
unavailable = time(0, 0)

# center instruction hours
instructionHours = {sun: [], \
                    mon: [time(15, 30), time(19, 30)], \
                    tue: [time(15, 30), time(19, 30)], \
                    wed: [time(15, 30), time(19, 30)], \
                    thu: [time(15, 30), time(19, 30)], \
                    fri: [], \
                    sat: [time(10, 00), time(14, 00)]}

# number of virtual instructors (move to config file)
virtual_instructors = 10


class Instructor:

    name = None
    schedule = None
    email_address = None
    availability = None
    availability_reported = None
    rank = None
    cost = 15
    max_hours_per_week = 20
    min_hours_per_day = 2
    work_day_preferences = {sun: 0, mon: 2, tue: 3, wed: 4, thu: 5, fri: 6, sat: 7}
    dayStrings = {sun: "Sunday", mon: "Monday", tue: "Tuesday", \
                           wed: "Wednesday", thu: "Thursday", fri: "Friday", \
                           sat: "Saturday"}
    isScheduled = False

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
        real_instructor.availability = {sun: (time(0, 0), time(0, 0)), \
                             mon: (row[5].value, row[6].value), \
                             tue: (row[8].value, row[9].value), \
                             wed: (row[11].value, row[12].value), \
                             thu: (row[14].value, row[15].value), \
                             fri: (time(0, 0), time(0, 0)), \
                             sat: (row[17].value, row[18].value)}

        # set rank
        real_instructor.rank = 999 #rank found
        #find row of the instructor, then assign rank from cell
        #cross references name in availability file to name in configuration sheet then assigns rank
        first_row = 2  # skip the headers
        last_row = Importer.config_ws.max_row
        last_col = Importer.config_ws.max_column
        for row in Importer.config_ws.iter_rows(min_row=first_row, max_col=last_col, max_row=last_row):
            #does my name match the name in the configuration file? If so, set rank.
            if real_instructor.name == row[0].value:
                real_instructor.rank = row[2].value

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
        return real_instructor

    @staticmethod
    def create_virtual_instructor(index=1):
        virtual_instructor = Instructor()
        virtual_instructor.schedule = {sun: [], mon: [], tue: [], wed: [], thu: [], fri: [], sat: []}
        virtual_instructor.availability_reported = datetime.now()
        virtual_instructor.email_address = "stafford@mathnasium.com" # Todo remove hard coded email address
        virtual_instructor.name = 'Gap '+ str(index)
        virtual_instructor.availability = instructionHours
        virtual_instructor.rank = 999
        virtual_instructor.cost = 15
        virtual_instructor.maxHrsPerWeek = 100
        virtual_instructor.minHrsPerDay = 0
        virtual_instructor.workDayPreferences = {sun: 0, mon: 2, tue: 3, wed: 4, thu: 5, fri: 6, sat: 7}
        virtual_instructor.dayStrings = {sun: "Sunday", mon: "Monday", tue: "Tuesday", \
                           wed: "Wednesday", thu: "Thursday", fri: "Friday", \
                           sat: "Saturday"}
        virtual_instructor.isScheduled = False
        return virtual_instructor

    @staticmethod
    def create_instructors(run_log):

        # Update Availability Work Sheet: When instructor not available, change none to time(0,0)
        first_row = 2  # skip the headers
        last_row = Importer.instructor_availability_ws.max_row
        last_col = Importer.instructor_availability_ws.max_column
        for row in Importer.instructor_availability_ws.iter_rows(min_row=first_row, max_col=last_col, max_row=last_row):
            for i in [5,6,8,9,11,12,14,15,17,18]:
                if row[i].value == None: row[i].value = time(0,0)
        for eachInstructor in instructors: print(eachInstructor.name)

        # Create Instructors
        Instructor.instructors = []
        print("\nCreating Instructors (from Instructor Availability and Configuration File")
        for row in Importer.instructor_availability_ws.iter_rows(min_row=first_row, max_col=last_col, max_row=last_row):
            Instructor.instructors.append(Instructor.create_real_instructor_from_row(run_log, row))
        # Add Virtual Instructors to identify periods that cannot covered by existing staff
        for i in range (1, virtual_instructors):
            Instructor.instructors.append(Instructor.create_virtual_instructor(i))
        for eachInstructor in Instructor.instructors: print("\tInstructor: ", eachInstructor.name,
                                                                     "Rank: ", eachInstructor.rank)
        return Instructor.instructors

    # define instructor sort
    def __lt__(self, other):
        return self.rank < other.rank

    def __str__(self):
        return str(self.name)

    def __print__(self):
        print(str(self.name, self.rank))

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
                # Bug fix: If no stop work ordered, stop work at end of day. Better fix check for no stop work
                if myStopTime == myStartTime: myStopTime = instructionHours[eachKey][1]
                self.schedule[eachKey] = [self.adjustTime(myStartTime), self.adjustTime(myStopTime)]
        return self.schedule

    def tuple(self, day):
        return [self.name, self.dayString(day), self.schedule[day][0], self.schedule[day][1]]
