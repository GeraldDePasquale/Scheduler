#test
import csv
import time
from datetime import datetime, date, timedelta, time

class Instructor:
    def __init__(self, csvObj, prev = None, next = None):

        # define days consistent with datetime.weekday()
        mon = 0
        tue = 1
        wed = 2
        thu = 3
        fri = 4
        sat = 5
        sun = 6

        self.prev = prev
        self.next = next

        # define instruction hours
        self.instructionHours = {sun: [], \
                                 mon: [time(15, 30), time(19, 30)], \
                                 tue: [time(15, 30), time(19, 30)], \
                                 wed: [time(15, 30), time(19, 30)], \
                                 thu: [time(15, 30), time(19, 30)], \
                                 fri: [], \
                                 sat: [time(10, 00), time(14, 00)]}

        # define work schedule
        self.schedule = {sun: [], mon: [], tue: [], wed: [], thu: [], fri: [], sat: []}

        # set instructor name
        self.name = csvObj[0]

        # set instructors hours of availability based upon the lines in the InstructorAvailability input file
        self.availability = {sun: ((time(int(csvObj[1].partition(':')[0]), int(csvObj[1].partition(':')[2], 0))), \
                                   (time(int(csvObj[2].partition(':')[0]), int(csvObj[2].partition(':')[2], 0)))), \
                             mon: ((time(int(csvObj[3].partition(':')[0]), int(csvObj[3].partition(':')[2], 0))), \
                                   (time(int(csvObj[4].partition(':')[0]), int(csvObj[4].partition(':')[2], 0)))), \
                             tue: ((time(int(csvObj[5].partition(':')[0]), int(csvObj[5].partition(':')[2], 0))), \
                                   (time(int(csvObj[6].partition(':')[0]), int(csvObj[6].partition(':')[2], 0)))), \
                             wed: ((time(int(csvObj[7].partition(':')[0]), int(csvObj[7].partition(':')[2], 0))), \
                                   (time(int(csvObj[8].partition(':')[0]), int(csvObj[8].partition(':')[2], 0)))), \
                             thu: ((time(int(csvObj[9].partition(':')[0]), int(csvObj[9].partition(':')[2], 0))), \
                                   (time(int(csvObj[10].partition(':')[0]), int(csvObj[10].partition(':')[2], 0)))), \
                             fri: ((time(int(csvObj[11].partition(':')[0]), int(csvObj[11].partition(':')[2], 0))), \
                                   (time(int(csvObj[12].partition(':')[0]), int(csvObj[12].partition(':')[2], 0)))), \
                             sat: ((time(int(csvObj[13].partition(':')[0]), int(csvObj[13].partition(':')[2], 0))), \
                                   (time(int(csvObj[14].partition(':')[0]), int(csvObj[14].partition(':')[2], 0))))}

        # set rank
        self.rank = int(csvObj[15])

        # set cost
        self.cost = float(csvObj[16])

        # set max hours per week
        self.maxHrsPerWeek = int(csvObj[17])

        # set min hours per day
        self.minHrsPerDay = int(csvObj[18])

        # sef work day preferences
        self.workDayPreferences = {sun: int(csvObj[19]), mon: int(csvObj[20]), tue: int(csvObj[21]), \
                                   wed: int(csvObj[22]), thu: int(csvObj[23]), fri: int(csvObj[24]), \
                                   sat: int(csvObj[25])}
        # set dayofweek strings
        self.dayStrings = {sun: "Sunday", mon: "Monday", tue: "Tuesday", \
                           wed: "Wednesday", thu: "Thursday", fri: "Friday", \
                           sat: "Saturday"}

        # set isSchedule flag - this variable is set by the scheduling process
        self.isScheduled = False

   # define instructor sort
    def __lt__(self, other):
        return self.rank < other.rank

    def __str__(self):
        return str(self.name)

    def __print__(self):
        print(str(self.name,self.rank))

    def prev(self):
       return self.prev

    def next(self):
        return self.next


    # weekdayString(self, dayofweek) returns string representation of the day
    def dayString(self, integer):
        return self.dayStrings[integer]

    # isAvailable returns true if the instructor is available to work at the time of the event
    def isAvailable(self, event):
        dayNeeded = event.eventTime.weekday()
        timeNeeded = event.eventTime.time()
        imScheduled = self.isScheduled
        availableStartTime = self.availability[dayNeeded][0]
        availableStopTime = self.availability[dayNeeded][1]
        result = (not imScheduled) and \
                 (timeNeeded >= availableStartTime) and \
                 (timeNeeded < availableStopTime)
        return result

    # isAvailableToOpen returns true if the instructor is available to start work when the center opens
    def isAvailableToOpen(self, event):
        dayNeeded = event.eventTime.weekday()
        timeNeeded = self.instructionHours[dayNeeded][0]
        imScheduled = self.isScheduled
        myStartTime = self.availability[dayNeeded][0]
        myStopTime = self.availability[dayNeeded][1]
        result = (not imScheduled) and \
                 (timeNeeded >= myStartTime) and \
                 (timeNeeded < myStopTime)
        return result

    # startWork adds the start work event time to the instructor's work schedule
    def startWork(self, event):
        dayNeeded = event.eventTime.weekday()
        timeNeeded = event.eventTime.time()
        self.schedule[dayNeeded].append(timeNeeded)
        self.isScheduled = True

    # startWorkWhenOpen adds start at open time to the instructor's work schedule
    def startWorkWhenOpen(self, event):
        dayNeeded = event.eventTime.weekday()
        timeNeeded = self.instructionHours[dayNeeded][0]
        self.schedule[dayNeeded].append(timeNeeded)
        self.isScheduled = True

    # stopWork adds the stop work event time to the instructor's work schedule
    def stopWork(self, event):
        self.schedule[event.eventTime.weekday()].append(event.eventTime.time())
        self.isScheduled = False

    # mustDepart returns true if the event time exceeds the instructors availability
    def mustDepart(self, event):
        dayNeeded = event.eventTime.weekday()
        timeNeeded = event.eventTime.time()
        myStopTime = self.availability[dayNeeded][1]
        result = (timeNeeded >= myStopTime)
        return result

    # departWork adds stop work event but leaves instructor in isScheduled state
    def departWork(self, event):
        self.schedule[event.eventTime.weekday()].append(event.eventTime.time())
        self.isScheduled = True

    # hoursScheduled returns the number of hours the employee is scheduled to work
    #    def hoursScheduledNow(self,event):


    # finalizeSchedule modifies the schedule so that each work day includes one start and one stop time
    def finalizeSchedule(self):
        for eachKey in self.schedule.keys():
            if self.schedule[eachKey]:
                myStartTime = max(self.schedule[eachKey][0],self.availability[eachKey][0])
                myStopTime = min(self.availability[eachKey][1],self.schedule[eachKey][len(self.schedule[eachKey])-1])
                #Bug fix: If no stop work ordered, stop work at end of day. Better way: check for non-existence of stop work
                if myStopTime == myStartTime: myStopTime = self.instructionHours[eachKey][1]
                self.schedule[eachKey] = [myStartTime,myStopTime]
#                if len(self.schedule[eachKey])%2 != 0:
#                    self.schedule[eachKey] = [myStartTime, myStopTime]
#                else:
#                    self.schedule[eachKey] = [myStartTime, self.schedule[eachKey][len(self.schedule[eachKey])-1]]
        return self.schedule
