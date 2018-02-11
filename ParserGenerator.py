from datetime import datetime

class ParserGenerator:

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

    def arrival_time_from_row(self, row, logfile):
        tokens = row[0].value.partition('/')
        month = int(tokens[0])
        tokens = tokens[2].partition('/')
        day = int(tokens[0])
        year = int(tokens[2])
        tokens = row[3].value.partition(':')
        hour = int(tokens[0])
        tokens = tokens[2].partition(' ')
        minute = int(tokens[0])
        meridiem = tokens[2]
        # Set 24 hour time based upon meridiem
        if ((meridiem == 'PM') and (hour != 12)): hour = hour + 12
        return datetime(year, month, day, hour, minute)

    def departure_time_from_row(self, row, logfile):
        tokens = row[0].value.partition('/')
        month = int(tokens[0])
        tokens = tokens[2].partition('/')
        day = int(tokens[0])
        year = int(tokens[2])
        tokens = row[4].value.partition(':')
        hour = int(tokens[0])
        tokens = tokens[2].partition(' ')
        minute = int(tokens[0])
        meridiem = tokens[2]
        # Set 24 hour time based upon meridiem
        if ((meridiem == 'PM') and (hour != 12)): hour = hour + 12
        return datetime(year, month, day, hour, minute)



   
