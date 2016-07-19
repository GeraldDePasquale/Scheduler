class InstructorScheduler:
    """A Mathnasium Instructor class"""

    defaultAvailability = []
    instructors = []

    def __init__ (self,csvFile):
        #Read Instructor Availability, Create Instructors
        headerline = True
        for each in csv.reader(csvFile):
            if headerline: headerline = False
            else:
                anInstructor = Instructor(each)
                self.instructors.append(anInstructor)
        self.addMinimumInstructors()

    def addInstructor(self, event):
        workAssigned = False
        for each in instructors:
            if not workAssigned and each.isAvailable(event.weekday()):
                each.startWork(event)
                workAssigned = true
        return workAssigned
                
    def dropInstructor(self, event):
        notDropped = True
