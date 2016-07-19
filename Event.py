import csv
import time
from datetime import datetime, date, timedelta
import math



class Event:

    """Mathnasium arrivals, departures, and cost"""
    events = []
    instructorChangeEvents = []
    peakEvents = []
    churnEvents = []
    
    def __init__(self, eventType=None, eventTime=None, studentObject=None, prev = None, next=None, churnTolerance=None):
        if (eventTime != None) and (eventType != None):
            self.student = studentObject.name
            self.grade = studentObject.grade
            self.eventTime = eventTime
            self.eventType = eventType
            self.studentObject = studentObject
            self.prev = prev
            self.next = next
            self.eventNumber = 0
            self.isArrivalEvent = eventType == 'Arrival'
            self.isDepartureEvent = eventType == 'Departure'
            self.instructorCount = 0 #instructor count after event
            self.studentCount = 0 #student count after the event
            self.isChurnEvent = False #set externally by event processor
            self.churnTolerance = 600 # 600 seconds (10 minutes)
            self.events.append(self)

    def __lt__(self,other):
        return self.eventTime < other.eventTime

    def __str__(self):
        return (str (self.eventNumber) + ',' + str(self.student) + ','
                + str(self.grade) + ',' + str(self.eventType) \
                + ',' + str(self.eventTime) + ',' + str(self.weekday()) \
                + ',' + str(self.studentCount) + ',' + str(self.studentInstructorRatio()) \
                + ',' + str(self.instructorCount) + '\n')
    def __print__(self):
        print(self.eventNumber, self.student, self.grade, self.eventType, \
              self.eventTime, self.weekday(), self.studentCount, \
              self.studentInstructorRatio(), self.instructorCount)

    def sort(self):
        return self.events.sort()

    def cost(self):
        if self.eventType == 'Arrival':
            return round(self.studentObject.cost(),4)
        elif self.eventType == 'Departure':
            return -round(self.studentObject.cost(),4)

    def weekday(self):
        daysOfWeek = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']
        aWeekday = daysOfWeek[self.eventTime.date().weekday()]
        return aWeekday

    def studentInstructorRatio(self):
        return round(self.studentCount/self.instructorCount, 1)

    def isDateChangeEvent(self):
        if self.prev is None: return True
        return self.prev.weekday() != self.weekday()

    def isInstructorChangeEvent(self):
        if self.prev is None:return True
        return self.prev.instructorCount != self.instructorCount
 
    def isPeakEvent(self):
        return self.isInstructorChangeEvent() and self.isArrivalEvent

    def isValleyEvent(self):
        return self.isInstructorChangeEvent() and self.isDepartureEvent
