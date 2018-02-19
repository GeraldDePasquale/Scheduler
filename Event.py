class Event:

    events = []
    instructorChangeEvents = []
    peakEvents = []
    churnEvents = []
    
    def __init__(self, event_type=None, event_time=None, student_object=None, prev = None, next=None, churn_tolerance=None):
        if (event_time != None) and (event_type != None):
            self.student = student_object.name
            self.grade = student_object.grade
            self.event_time = event_time
            self.event_type = event_type
            self.student_object = student_object
            self.prev = prev
            self.next = next
            self.event_number = 0
            self.is_arrival_event = event_type == 'Arrival'
            self.is_departure_event = event_type == 'Departure'
            self.instructor_count = 0 #instructor count after event
            self.student_count = 0 #student count after the event
            self.is_churn_event = False #set externally by event processor
            self.churn_tolerance = 600 # 600 seconds (10 minutes) # Todo move to config file
            self.events.append(self)

    def __lt__(self, other):
        return self.event_time < other.event_time

    def __str__(self):
        return (str (self.event_number) + ',' + str(self.student) + ','
                + str(self.grade) + ',' + str(self.event_type) \
                + ',' + str(self.event_time) + ',' + str(self.weekday()) \
                + ',' + str(self.student_count) + ',' + str(self.student_instructor_ratio()) \
                + ',' + str(self.instructor_count) + '\n')
    def __print__(self):
        print(self.event_number, self.student, self.grade, self.event_type, \
              self.event_time, self.weekday(), self.student_count, \
              self.student_instructor_ratio(), self.instructor_count)

    def tuple(self):
        return [self.event_number, self.student, self.grade, self.event_type,
                self.event_time, self.weekday(), self.student_count,
                self.student_instructor_ratio(), self.instructor_count]

    def sort(self):
        return self.events.sort()

    def cost(self):
        if self.event_type == 'Arrival':
            return round(self.student_object.cost, 4)
        elif self.event_type == 'Departure':
            return -round(self.student_object.cost, 4)

    def weekday(self):
        daysOfWeek = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']
        aWeekday = daysOfWeek[self.event_time.date().weekday()]
        return aWeekday

    def student_instructor_ratio(self):
        return round(self.student_count / self.instructor_count, 1)

    def is_date_change_event(self):
        if self.prev is None: return True
        return self.prev.weekday() != self.weekday()

    def is_instructor_change_event(self):
        if self.prev is None:return True
        return self.prev.instructor_count != self.instructor_count
 
    def is_peak_event(self):
        return self.is_instructor_change_event() and self.is_arrival_event

    def is_valley_event(self):
        return self.is_instructor_change_event() and self.is_departure_event

    def is_summary_event(self):
        return (not self.is_churn_event and (self.is_date_change_event() or self.is_instructor_change_event()))
