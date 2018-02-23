from datetime import timedelta
from ParserGenerator import ParserGenerator


class Student:

    students = []
    log_file = None
    config_ws = None
    name = None
    sessionTime = 1  # 1 hour
    lowCost = 1 / 5  # 1 to 5 student ratio (Grades 2 .. 5)
    mediumCost = 1 / 4  # 1 to 4 student ratio (Grades 6 .. 8)
    highCost = 1 / 3  # 1 to 3 student ratio (Grades 0, 1, 9 .. 13+)
    veryHighCost = 1 / 1  # Private Lesson
    grade = None
    arrival_time = None
    departure_time = None
    grade_dictionary = {}
    cost = None

    def __lt__(self, other):
        return self.arrival_time < other.arrivalTime

    @staticmethod
    def get_cost(grade): # Todo move costs and grades to config
        if grade in []: return Student.highCost
        elif grade in ['2', '3', '4', '5', '6', '7', '8']: return Student.highCost
        elif grade in ['', 'K', '1', '2', '9', '10', '11', '12', '15', 'U']: return Student.highCost

    @staticmethod
    def new_from_row(row, run_log_ws):
        instance = Student()
        instance.name = row[1].value + ' ' + row[2].value
        try:
            # get datetime when student arrives at center
            instance.arrival_time = ParserGenerator.arrival_time_from_row(row, run_log_ws)
        except:
            run_log_ws.write(str(row) + '\n')
        try:
            # get datetime when student departs from center
            instance.departure_time = ParserGenerator.departure_time_from_row(row, run_log_ws)
        except:
            # if the departure time is absent (empty string) set departure to arrival time + session time
            instance.departure_time = instance.arrival_time + timedelta(hours=instance.sessionTime)
            # Adjust session start and stop time if necessary
        # set grade
        instance.grade = Student.grade_dictionary.get(instance.name, "U")
        instance.cost = Student.get_cost(str(instance.grade))
#        print("Name: " + instance.name + " Grade: " + str(instance.grade))
        return instance

    def create_grade_dictionary_from_ws(self, student_data_ws):
        my_grade_dictionary = {}
        first_row = 2  # skip the headers
        last_row = student_data_ws.max_row
        last_col = student_data_ws.max_column
        for row in student_data_ws.iter_rows(min_row=first_row, max_col=last_col, max_row=last_row):
            my_grade_dictionary[row[1].value] = row[6].value
        Student.grade_dictionary = my_grade_dictionary
        return self.grade_dictionary

    @staticmethod
    def initialize_students(attendance_ws, student_data_ws, run_log_ws):
        Student().create_grade_dictionary_from_ws(student_data_ws)
        print("\nCreating students from Student Attendance Report")
        first_row = 2  # skip the headers
        last_row = attendance_ws.max_row
        last_col = attendance_ws.max_column
        for row in attendance_ws.iter_rows(min_row=first_row, max_col=last_col, max_row=last_row):
            Student.students.append(Student.new_from_row(row, run_log_ws))
        return Student.students
