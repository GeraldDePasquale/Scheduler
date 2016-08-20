import ParserGenerator
from datetime import datetime, date, timedelta
from ParserGenerator import ParserGenerator

class Student:
    """A Mathnasium Student class"""
    lowCost = 1/5 #1 to 5 student ratio (Grades 2 .. 5)
    mediumCost = 1/4 #1 to 4 student ratio (Grades 6 .. 8)
    highCost = 1/3 #1 to 3 student ratio (Grades 0, 1, 9 .. 13+)
    veryHighCost = 1/1 #Private Lesson
    sessionTime = 1 #hour
  
    def __init__(self, line, logFile):
        #Get name, grade, arrival, departure strings
        self.name = line[0] #first and last name
        try:
            self.arrivalTime = ParserGenerator().pgDatetime(line[2], logFile) #datetime when arrives at center
        except:
            logFile.write(line + ',' + 'Invalid Arrival Time: ' + line[2] + '\n')
        try:
            self.departureTime = ParserGenerator().pgDatetime(line[3], logFile) #datetime when departs from center
        except:
            #if the departime time is absent (empty string) set departure to arrival time + session time
            self.departureTime = self.arrivalTime + timedelta(hours = self.sessionTime)
        if line[1] not in ['K', '1', '2', '3', '4', '5', '6','7', '8','9','10','11','12']:
            logFile.write(str(self.name) + ',' + str(self.arrivalTime) + ',' + 'No Grade\n')
            self.grade = 'U' #undefined
        else: self.grade =  self.grade = line[1]

    def __lt__(self,other):
        return self.arrivalTime < other.arrivalTime

    def cost(self):
        if self.grade in []:
            return self.highCost
        elif self.grade in ['3', '4', '5','6', '7', '8']:
            return self.highCost
        elif self.grade in ['','K', '1', '2', '9','10','11','12','13', 'U']:
            return self.highCost
