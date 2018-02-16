# Mathnasium Staffing Forecast
#   Forecasts instructor staffing requirements for scheduling purposes.
#   Forecasts based upon a student arrival-departure event model.

# Outputs:
# Excel Workbook containing:
# 1. Instructor Schedules
# 2. Summary Staffing Forecast Sheet provides number of instructors required each instruction day. Shows only
#   events when actual increases or decreases in staffing levels are required.
# 3. Detailed Stafford Forecast Sheet provides same data as Summary Staffing Forecast except that it provides
#    every event used for forcasting.

# Inputs:
# 1. Radius Attendance Report
# 2. Radius Student Report
# 3. Configuration File - includes instructor ranking for scheduling purposes. Also provides other instructor
# specific data (e.g. cost / hour)
# 4. Work Availability Spreadsheet - provides instructor availability as input using Google Form: Work Availability

# Hardcoded Parameters:
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
import os
import os.path
from datetime import datetime
from tkinter import Tk, filedialog, simpledialog
from openpyxl import Workbook, worksheet, load_workbook
import Event
import Student
from Event import Event
from Instructor import Instructor
from Student import Student
from Importer import Importer
from Reporter import Reporter


def main():

    root = Tk()
    root.withdraw()
    run_time = datetime.now().strftime('%Y%m%d%H%M') # used for file extensions, makes sorting easy
    print("Mathnasium Scheduler Starts")
    default_directory = "C:\\ProgramData\\MathnasiumScheduler"
    FILEOPENOPTIONS = dict(defaultextension='.csv', filetypes=[('XLSX', '*.xlsx'), ('CSV file', '*.csv')])
    # Todo-jerry add center name picklist, get center names from configuraton file
    center_name = simpledialog.askstring("Name prompt", "Enter Center Name")
    scheduling_data_sheets = Importer(run_time, default_directory, center_name, FILEOPENOPTIONS) #load the working data

    #Create Schedule Workbook
    # Create Run Workbook
    run_wb_path = default_directory + "\\" + center_name + "." + run_time + ".xlsx"
    run_wb = Workbook()
    run_wb.save(run_wb_path)
    schedule_by_name_ws = run_wb.create_sheet("Schedule By Name", index=0)
    schedule_by_day_ws = run_wb.create_sheet("Schedule By Day", index=1)
    forecast_summary_ws = run_wb.create_sheet("Summary Forecast", index=2)
    forecast_detailed_ws = run_wb.create_sheet("Detailed Forecast", index=3)
    run_log_ws = run_wb.create_sheet("Runlog", index=4)
    #ToDo Write run_log to run_log_ws
    run_log = []

    # Open Log File
#    directory_warnings = default_directory + "\\Warnings"
#   if not os.path.exists(directory_warnings): os.makedirs(directory_warnings)
#    logfile_name = directory_warnings + "\\Forecast Warnings.csv"
#    run_log_ws = open(logfile_name, 'w')
#    print("Opened ", logfile_name)

    Instructor.initialize(scheduling_data_sheets.instructor_availability_ws,
                          scheduling_data_sheets.config_ws,
                          run_log)

    #Create Students
    print("Creating students from Student Attendance Report\n")
    students = []
    student_ws = scheduling_data_sheets.attendance_ws  #attendance_wb.active
    first_row = 2 #skip the headers
    last_row = student_ws.max_row
    last_col = student_ws.max_column
    for row in student_ws.iter_rows(min_row = first_row, max_col = last_col, max_row=last_row):
        students.append(Student(row, run_log))

    #Create Events
    print("Creating events from student arrivals and departures\n")
    events = []
    for each_student in students:
        events.append(Event('Arrival', each_student.arrivalTime, each_student))
        events.append(Event('Departure', each_student.departureTime, each_student))
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
    event_groups = {sun: [], mon: [], tue: [], wed: [], thu: [], fri: [], sat: []}
    for each_event in events:
        event_groups[each_event.eventTime.weekday()].append(each_event)

    print("\tDetermining cost of each day")
    costsOfEventGroups = {}
    for each_day in event_groups.keys():
        cost = 0.0
        for each_event in event_groups[each_day]:
            if each_event.isArrivalEvent: cost = cost + each_event.cost()
        costsOfEventGroups[each_day] = round(cost, 1)
        print("\t\tDay: ", each_day, "Cost: ", costsOfEventGroups[each_day])

    # Sort and process each group of events
    for each_day in event_groups.keys():
        print("\n\tProcessing Day: ", str(each_day))
        event_groups[each_day].sort()
        instructorsMinimum = 2.0  # the minimum staffing level
        instructorsRequired = 0.0  # actual number of instructors required to meet student demand
        eventnumber = 1  # first event number
        studentcount = 0  # start with zero students
        for each_event in event_groups[each_day]:
            # Set event number
            each_event.eventNumber = eventnumber
            # Set event's previous and next events
            if eventnumber != 1: each_event.prev = events[events.index(each_event) - 1]
            if eventnumber != len(event_groups[each_day]): each_event.next = event_groups[each_day][
                event_groups[each_day].index(each_event) + 1]
            eventnumber = eventnumber + 1  # next event number
            # Maintain student count
            if (each_event.isArrivalEvent):
                studentcount = studentcount + 1
            elif (each_event.isDepartureEvent):
                studentcount = studentcount - 1
            each_event.studentCount = studentcount
            # Compute/maintain the actual number of instructors required
            instructorsRequired = instructorsRequired + each_event.cost()
            # Compute/maintain the number of instructors to staff (minimum is instructorsMinimum)
            each_event.instructorCount = max(instructorsMinimum, math.ceil(instructorsRequired))

        print("\t\tCollecting Instructor Change Events")
        instructor_change_events = []
        for each_event in event_groups[each_day]:
            if each_event.isInstructorChangeEvent():
                instructor_change_events.append(each_event)

        print("\t\tMarking Churn Events")
        tolerance = 360  # seconds (6 minutes)
        for i in range(len(instructor_change_events) - 1):
            event = instructor_change_events[i]
            nextEvent = instructor_change_events[i + 1]
            event.isChurnEvent = event.isPeakEvent() and nextEvent.isValleyEvent() \
                                 and (nextEvent.eventTime - event.eventTime).seconds < tolerance

        print("\t\tScheduling Instructors")
        instructors = Instructor.instructors
        instructors.sort()
        unscheduled_instructors = instructors
        scheduled_instructors = []
        dateChangeEvents = 0

        for each_event in event_groups[each_day]:
            if each_event.isDateChangeEvent():
                dateChangeEvents = dateChangeEvents + 1
                # Schedule minimum number of instructors needed to open
                count_scheduled = 0
                unscheduled_instructors.sort()
                instructors_changed = []
                for this_instructor in unscheduled_instructors:
                    if this_instructor.isAvailableToOpen(each_event) and (count_scheduled < instructorsMinimum):
                        count_scheduled = count_scheduled + 1
                        this_instructor.startWorkWhenOpen(each_event)
                        scheduled_instructors.append(this_instructor)
                        # save pointers to instructors for removal from unscheduled list
                        instructors_changed.append(this_instructor)
                for i in instructors_changed: unscheduled_instructors.remove(i)

            # Check for and remove departing Instructors
            departed_instructors = []
            for this_instructor in scheduled_instructors:
                departed = False
                if this_instructor.mustDepart(each_event):
                    this_instructor.departWork(each_event)
                    departed_instructors.append(this_instructor)

            instructor_change_needed = each_event.instructorCount - len(scheduled_instructors)
            if not each_event.isChurnEvent and instructor_change_needed > 0:
                # Schedule available instructors
#                print("\t\t\tSchedule Instructor")
                count_scheduled = 0
                unscheduled_instructors.sort()
#                print("Unscheduled Instructors: " + str(len(unscheduled_instructors)))
                while (count_scheduled < instructor_change_needed):
                    # Find instructor and schedule instructor
                    instructors_changed = []
                    for this_instructor in unscheduled_instructors:
                        if this_instructor.isAvailable(each_event) and (count_scheduled < instructor_change_needed):
                            count_scheduled = count_scheduled + 1
                            this_instructor.startWork(each_event)
                            scheduled_instructors.append(this_instructor)
                            # Save pointers to newly scheduled instructors for removal from unscheduled list
                            instructors_changed.append(this_instructor)
                # Remove newly scheduled instructors from unscheduled list
                for i in instructors_changed: unscheduled_instructors.remove(i)

            if not each_event.isChurnEvent and instructor_change_needed < 0:
                # Unschedule instructors
#                print("\t\t\tUnschedule Instructor")
                count_unscheduled = 0
                instructors_changed = []
                unscheduled_instructors.sort()
                while count_unscheduled < abs(instructor_change_needed):
                    for this_instructor in reversed(scheduled_instructors):
                          if count_unscheduled < abs(instructor_change_needed):
                            count_unscheduled = count_unscheduled + 1
                            this_instructor.stopWork(each_event)
                            unscheduled_instructors.append(this_instructor)
                            instructors_changed.append(this_instructor)
                for i in instructors_changed: scheduled_instructors.remove(i)

            # Finalize schedules after last departure or final event of the day
            if (each_event == event_groups[each_day][len(event_groups[each_day]) - 1]) or each_event.next.isDateChangeEvent():
                instructors_changed = []
                for this_instructor in scheduled_instructors:
                    this_instructor.isScheduled = False
                    unscheduled_instructors.append(this_instructor)
                    instructors_changed.append(this_instructor)
                for i in instructors_changed: scheduled_instructors.remove(i)
                for i in unscheduled_instructors: i.finalizeSchedule()

    Reporter().write_all(events, instructors, forecast_detailed_ws, forecast_summary_ws, schedule_by_name_ws)

    print("Closing files")
    scheduling_data_sheets.close_workbooks()
    run_wb.save(run_wb_path)
    run_wb.close()
    print("Launching Excel")
    os.system("start excel " + run_wb_path )
    print("Done")

if __name__ == '__main__': main()
