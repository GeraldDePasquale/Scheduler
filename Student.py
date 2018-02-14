import ParserGenerator
from datetime import datetime, date, timedelta
from ParserGenerator import ParserGenerator


class Student:
    """A Mathnasium Student class"""
    lowCost = 1 / 5  # 1 to 5 student ratio (Grades 2 .. 5)
    mediumCost = 1 / 4  # 1 to 4 student ratio (Grades 6 .. 8)
    highCost = 1 / 3  # 1 to 3 student ratio (Grades 0, 1, 9 .. 13+)
    veryHighCost = 1 / 1  # Private Lesson
    sessionTime = 1  # hour

    def __init__(self, row, run_log_ws):
        self.name = row[1].value + ' ' + row[2].value
        try:
            # get datetime when student arrives at center
            self.arrivalTime = ParserGenerator().arrival_time_from_row(row, run_log_ws)
        except:
            run_log_ws.write(str(row) + '\n')
        try:
            # get datetime when student departs from center
            self.departureTime = ParserGenerator().departure_time_from_row(row, run_log_ws)
        except:
            # if the departure time is absent (empty string) set departure to arrival time + session time
            self.departureTime = self.arrivalTime + timedelta(hours=self.sessionTime)
            # Adjust session start and stop time if necessary
        self.grade = 'U'

    def __lt__(self, other):
        return self.arrivalTime < other.arrivalTime

    def cost(self):
        if self.grade in []:
            return self.highCost
        elif self.grade in ['3', '4', '5', '6', '7', '8']:
            return self.highCost
        elif self.grade in ['', 'K', '1', '2', '9', '10', '11', '12', '13', 'U']:
            return self.highCost
