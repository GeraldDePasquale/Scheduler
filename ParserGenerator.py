import csv
import time
from datetime import datetime, date, timedelta, time

class ParserGenerator:

#    def __init__(self):
    
    def pgDatetime(self, parseString, logFile):
    #parse the string
        tokens = parseString.partition('/')
        month = int(tokens[0])
        tokens = tokens[2].partition('/')
        day = int(tokens[0])
        tokens = tokens[2].partition(' ')
        year = int(tokens[0])
        tokens = tokens[2].partition(':')
        hour = int(tokens[0])
        tokens = tokens[2].partition(' ')
        minute = int(tokens[0])
        meridiem = tokens[2]
        #Set 24 hour time based upon meridiem
        if ((meridiem == 'PM') and (hour != 12)): hour = hour + 12
        return datetime(year,month,day,hour,minute)


   
