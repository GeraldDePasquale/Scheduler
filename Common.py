import csv
import time
from datetime import datetime, date, timedelta, time

class Common:
    def __init__(self):
    #define days consistent with datetime.weekday()
    self.mon = 0
    self.tue = 1
    self.wed = 2
    self.thu = 3
    self.fri = 4
    self.sat = 5
    self.sun = 6

    #set dayofweek strings
    self.dayStrings = {sun: "Sunday", mon: "Monday", tue: "Tuesday",\
                       wed: "Wednesday", thu: "Thursday", fri: "Friday",\
                       sat: "Saturday"}
    
