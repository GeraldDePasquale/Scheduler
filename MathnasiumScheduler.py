# Mathnasium Staffing Forecast
#   Forecasts instructor staffing requirements for scheduling purposes.
#   It forecasts based upon a student arrival-departure event model.
#   OOArrivals.5.py March 5, 2015 Linked list modification
#   OOArrivals.6.py March 5, 2015 Use Event class variables

# Outputs:
#  (1) Summary Staffing Requirements (Instructor Forecast - Summary.csv)
#   number of instructors required each instruction day. Shows only
#   times when increases and decreases in staffing levels are required.
#  (2) Detailed Staffing Requirements (Instructor Forecast - Detailed.csv)
#   number of instructors required each instruction day. Shows every
#   event used for forcasting.
#  (3) Forecast Noticesg (Forecast Warnings.csv)
#   Notices and Warnings the forecast that was run (e.g. defaults used
#   during the run, data errors, assumptions used to handle data errors). 

# Inputs:
#  (1) Custom M2 Attendance Report (Attendance Report.csv) for the period
#   of time, used to forecast staffing requirements exported in csv format
#  (2) Input parameters
#    lowCost = 1/5 instructor to student ratio (Grades 2 .. 5)
#    mediumCost = 1/4 instructor to student ratio (Grades 6 .. 8)
#    highCost = 1/3 instructor to student ratio (Grades 0, 1, 9+)
#    veryHighCost = 1.0 private tutoring
#    churnTolerance = 360
#       The number of seconds per hour that will be tolerated to minimize
#       churn and avoid overstaffing due to students arrivals and departures.
#       For instance, if we need desire to maintain a 1:3 instructor:student
#       ratio and we have 10:31, how long will we tolerate it before
#       we bring in another instructor. Due to student churn, that ratio may
#       return to 10:30 in a matter of minutes. How many minutes are tolerable?
#       Default: 6 minutes per hour.


import csv
import math
import os.path
from datetime import datetime
from tkinter import Tk, filedialog, simpledialog
from openpyxl import workbook, worksheet, load_workbook
import Event
import Student
from Event import Event
from Instructor import Instructor
from Student import Student

def main():

    root = Tk()
    root.withdraw()

    print("MathnasiumScheduler Starts")

    directory = "C:\\ProgramData\\MathnasiumScheduler"
    FILEOPENOPTIONS = dict(defaultextension='.csv', filetypes=[('XLSX', '*.xlsx'), ('CSV file', '*.csv')])

    #Get Center Name
    centerName = simpledialog.askstring("Name prompt", "Enter Center Name")

    #Open Attendance File
    directory = "C:\\ProgramData\\MathnasiumScheduler"
    student_attendance_report_file = filedialog.askopenfilename(parent=root, initialdir=directory,
                                                          title='Select Student Attendance Report', **FILEOPENOPTIONS)
    student_attendance_wb = load_workbook(student_attendance_report_file, data_only=True, guess_types=True)
    print("Opened", student_attendance_report_file)

    # Open Log File
    directoryWarnings = directory + "\\Warnings"
    if not os.path.exists(directoryWarnings):
        os.makedirs(directoryWarnings)
    logFileName = directoryWarnings + "\\Forecast Warnings.csv"
    logFile = open(logFileName, 'w')
    print("Opened ", logFileName)

    Instructor.initialize(logFile, root)

    #Open Instructor Schedule File
    prefix = datetime.now().strftime('%Y%m%d%H%M%S')
    instructorScheduleFileName = directory+"\\"+centerName+" Instructor Schedule."+prefix+".csv"
    instructorScheduleFile = open(instructorScheduleFileName, 'w')
    print("Opened ", instructorScheduleFileName)

    #Create Students
    print("Creating students from Student Attendance Report\n")
    students = []
    student_ws = student_attendance_wb.active
    first_row = 2 #skip the headers
    last_row = student_ws.max_row
    last_col = student_ws.max_column
    for row in student_ws.iter_rows(min_row = first_row, max_col = last_col, max_row=last_row):
        students.append(Student(row,logFile))

    #Create Events
    print("Creating events from student arrivals and departures\n")
    events = []
    for eachStudent in students:
        events.append(Event('Arrival', eachStudent.arrivalTime, eachStudent))
        events.append(Event('Departure', eachStudent.departureTime, eachStudent))
    events.sort()

    #Executing Events
    print("Executing events and collecting information")
    # Gather events by week day
    # define days consistent with datetime.weekday() - - - there has to be a better way than this
    mon = 0
    tue = 1
    wed = 2
    thu = 3
    fri = 4
    sat = 5
    sun = 6

    # Group events by day
    eventgroups = {sun: [], mon: [], tue: [], wed: [], thu: [], fri: [], sat: []}
    for eachEvent in events:
        eventgroups[eachEvent.eventTime.weekday()].append(eachEvent)

    print("\tDetermining cost of each day")
    costsOfEventGroups = {}
    for eachDay in eventgroups.keys():
        cost = 0.0
        for eachEvent in eventgroups[eachDay]:
            if eachEvent.isArrivalEvent: cost = cost + eachEvent.cost()
        costsOfEventGroups[eachDay] = round(cost, 1)
        print("\t\tDay: ", eachDay, "Cost: ", costsOfEventGroups[eachDay])

    # Sort and process each group of events
    for eachDay in eventgroups.keys():
        print("\n\tProcessing Day: ", str(eachDay))
        eventgroups[eachDay].sort()
        instructorsMinimum = 2.0  # the minimum staffing level
        instructorsRequired = 0.0  # actual number of instructors required to meet student demand
        eventnumber = 1  # first event number
        studentcount = 0  # start with zero students
        for eachEvent in eventgroups[eachDay]:
            # Set event number
            eachEvent.eventNumber = eventnumber
            # Set event's previous and next events
            if eventnumber != 1: eachEvent.prev = events[events.index(eachEvent) - 1]
            if eventnumber != len(eventgroups[eachDay]): eachEvent.next = eventgroups[eachDay][
                eventgroups[eachDay].index(eachEvent) + 1]
            eventnumber = eventnumber + 1  # next event number
            # Maintain student count
            if (eachEvent.isArrivalEvent):
                studentcount = studentcount + 1
            elif (eachEvent.isDepartureEvent):
                studentcount = studentcount - 1
            eachEvent.studentCount = studentcount
            # Compute/maintain the actual number of instructors required
            instructorsRequired = instructorsRequired + eachEvent.cost()
            # Compute/maintain the number of instructors to staff (minimum is instructorsMinimum)
            eachEvent.instructorCount = max(instructorsMinimum, math.ceil(instructorsRequired))

        print("\t\tCollecting Instructor Change Events")
        instructorChangeEvents = []
        for eachEvent in eventgroups[eachDay]:
            if eachEvent.isInstructorChangeEvent():
                instructorChangeEvents.append(eachEvent)

        print("\t\tMarking Churn Events")
        tolerance = 360  # seconds (6 minutes)
        for i in range(len(instructorChangeEvents) - 1):
            event = instructorChangeEvents[i]
            nextEvent = instructorChangeEvents[i + 1]
            event.isChurnEvent = event.isPeakEvent() and nextEvent.isValleyEvent() \
                                 and (nextEvent.eventTime - event.eventTime).seconds < tolerance

        print("\t\tScheduling Instructors")
        instructors = Instructor.instructors
        instructors.sort()
        unscheduledInstructors = instructors
        scheduledInstructors = []
        dateChangeEvents = 0

        for eachEvent in eventgroups[eachDay]:
            if eachEvent.isDateChangeEvent():
                dateChangeEvents = dateChangeEvents + 1
                # Schedule minimum number of instructors needed to open
                countScheduled = 0
                unscheduledInstructors.sort()
                instructorsChanged = []
                for thisInstructor in unscheduledInstructors:
                    if thisInstructor.isAvailableToOpen(eachEvent) and (countScheduled < instructorsMinimum):
                        countScheduled = countScheduled + 1
                        thisInstructor.startWorkWhenOpen(eachEvent)
                        scheduledInstructors.append(thisInstructor)
                        # save pointers to instructors for removal from unscheduled list
                        instructorsChanged.append(thisInstructor)
                for i in instructorsChanged: unscheduledInstructors.remove(i)

            # Check for and remove departing Instructors
            departedInstructors = []
            for thisInstructor in scheduledInstructors:
                departed = False
                if thisInstructor.mustDepart(eachEvent):
                    thisInstructor.departWork(eachEvent)
                    departedInstructors.append(thisInstructor)

            instructorChangeNeeded = eachEvent.instructorCount - len(scheduledInstructors)
            if not eachEvent.isChurnEvent and instructorChangeNeeded > 0:
                # Schedule available instructors
#                print("\t\t\tSchedule Instructor")
                countScheduled = 0
                unscheduledInstructors.sort()
#                print("Unscheduled Instructors: " + str(len(unscheduledInstructors)))
                while (countScheduled < instructorChangeNeeded):
                    # Find instructor and schedule instructor
                    instructorsChanged = []
                    for thisInstructor in unscheduledInstructors:
                        if thisInstructor.isAvailable(eachEvent) and (countScheduled < instructorChangeNeeded):
                            countScheduled = countScheduled + 1
                            thisInstructor.startWork(eachEvent)
                            scheduledInstructors.append(thisInstructor)
                            # Save pointers to newly scheduled instructors for removal from unscheduled list
                            instructorsChanged.append(thisInstructor)
                # Remove newly scheduled instructors from unscheduled list
                for i in instructorsChanged: unscheduledInstructors.remove(i)

            if not eachEvent.isChurnEvent and instructorChangeNeeded < 0:
                # Unschedule instructors
#                print("\t\t\tUnschedule Instructor")
                countUnscheduled = 0
                instructorsChanged = []
                unscheduledInstructors.sort()
                while countUnscheduled < abs(instructorChangeNeeded):
                    for thisInstructor in reversed(scheduledInstructors):
                          if countUnscheduled < abs(instructorChangeNeeded):
                            countUnscheduled = countUnscheduled + 1
                            thisInstructor.stopWork(eachEvent)
                            unscheduledInstructors.append(thisInstructor)
                            instructorsChanged.append(thisInstructor)
                for i in instructorsChanged: scheduledInstructors.remove(i)

            # Finalize schedules after last departure or final event of the day
            if (eachEvent == eventgroups[eachDay][len(eventgroups[eachDay]) - 1]) or eachEvent.next.isDateChangeEvent():
                instructorsChanged = []
                for thisInstructor in scheduledInstructors:
                    thisInstructor.isScheduled = False
                    unscheduledInstructors.append(thisInstructor)
                    instructorsChanged.append(thisInstructor)
                for i in instructorsChanged: scheduledInstructors.remove(i)
                for i in unscheduledInstructors: i.finalizeSchedule()

    print("\n Writing ......")
    print("\t C:\\Users\\Gerald\\Downloads\\Scheduling\\"+centerName+" Instructor Forecast - Summary.csv")
    print("\t C:\\Users\\Gerald\\Downloads\\Scheduling\\"+centerName+" Instructor Forecast - Detailed.csv")
    print("\t C:\\Users\\Gerald\Downloads\\Scheduling\\"+centerName+" Instructor Schedule.csv\n")
    # Write Attendance Forecasts
    summaryForecastFile = open(directory+'\\'+centerName+' Instructor Forecast - Summary.csv', 'w')
    detailedForecastFile = open(directory+'\\'+centerName+' Instructor Forecast - Detailed.csv', 'w')
    summaryForecastFile.write(
        str('Event #,Student Name,Grade,Event,Time,Day,Student Count,Student:Instructor,Instructors Required\n'))
    detailedForecastFile.write(
        str('Event #,Student Name,Grade,Event,Time,Day,Student Count,Student:Instructor,Instructors Required\n'))
    for each in events:
        # Write Summary Forecast
        if ((not each.isChurnEvent) and (each.isDateChangeEvent() or each.isInstructorChangeEvent())):
            summaryForecastFile.write(str(each))
        # Write Detailed Forecast
        detailedForecastFile.write(str(each))

    # Write Instructor Schedules
    instructorScheduleFile.write("Instructor Name," + "Day," + "Start Time," + "Stop Time\n")
    for eachInstructor in instructors:
        for eachDay in eachInstructor.schedule.keys():
            if eachInstructor.schedule[eachDay]:
                instructorScheduleFile.write(str(eachInstructor.name) + "," + eachInstructor.dayString(eachDay) + "," \
                                             + str(eachInstructor.schedule[eachDay][0]) + "," \
                                             + str(eachInstructor.schedule[eachDay][1]) + "\n")
    print("Closing all files\n")
    student_attendance_wb.close()
#    instructorAvailabilityFile.close()
    summaryForecastFile.close()
    detailedForecastFile.close()
    instructorScheduleFile.close()
    logFile.close()

    print("Done")


if __name__ == '__main__': main()
