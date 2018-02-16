import time
from datetime import datetime, time
from tkinter import Tk, filedialog, simpledialog
from openpyxl import Workbook, worksheet, load_workbook

root = Tk()
root.withdraw()
root_directory = None
# Set Default Directory
# Todo-jerry remove hard coded directories
attendance_wb = None
student_data_wb = None
config_wb = None
instructor_availability_wb = None
working_copy = "Working Copy"
attendance_ws = None
student_data_ws = None
config_ws = None
instructor_availability_ws = None



class Importer:

    def __init__(self, run_time, directory, center_name, FILEOPENOPTIONS):
        # Load Attendance Report
        # Source: Radius>>Reports>>Student Attendance Type: .xlsx
        student_attendance_file_name = filedialog.askopenfilename(parent=root, initialdir=directory,
                                                                  title='Select Radius Student Attendance Report',
                                                                  **FILEOPENOPTIONS)
        attendance_wb = load_workbook(student_attendance_file_name, data_only=True, guess_types=True)
        Importer.attendance_wb = attendance_wb
        Importer.attendance_ws = attendance_wb.active  # working sheet
        Importer.attendance_ws.title = "Attendance"
        print("Imported Attendance Report", student_attendance_file_name)

        # Load Student Report
        # Source: Radius>>Reports>>Student  Type: .xlsx
        student_data_file_name = filedialog.askopenfilename(parent=root, initialdir=directory,
                                                            title='Select Student Report',
                                                            **FILEOPENOPTIONS)
        student_data_wb = load_workbook(student_data_file_name, data_only=True, guess_types=True)
        Importer.student_data_wb = student_data_wb
        Importer.student_data_ws = student_data_wb.active  # working sheet
        Importer.student_data_ws.title = "Students"
        print("Imported Student Report", student_data_file_name)

        # Load Configuration
        config_file_name = filedialog.askopenfilename(parent=root, initialdir=directory,
                                                      title='Select Configuration File',
                                                      **FILEOPENOPTIONS)
        config_wb = load_workbook(config_file_name, data_only=True, guess_types=True)
        Importer.config_wb = config_wb
        Importer.config_ws = config_wb.active
        print("Loaded ", config_file_name)

        # Load Instructor Availabilty
        instructor_availability_file_name = filedialog.askopenfilename(parent=root, initialdir=directory,
                                                                       title='Select Instructor Availability File',
                                                                       **FILEOPENOPTIONS)
        instructor_availability_wb = load_workbook(instructor_availability_file_name, data_only=True, guess_types=True)
        Importer.instructor_availability_wb = instructor_availability_wb
        Importer.instructor_availability_ws = instructor_availability_wb.active
        print("Loaded ", instructor_availability_file_name)

    def close_workbooks(self):
        Importer.attendance_wb.close()
        Importer.student_data_wb.close()
        Importer.config_wb.close()
        Importer.instructor_availability_wb.close()
