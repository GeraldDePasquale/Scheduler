

class Reporter:
    forecast_headers = ["Event #", "Student Name", "Grade", "Event", "Time", "Day",
                        "Student Count", "Student:Instructor", "Instructors Required"]
    schedule_headers = ["Instructor Name", "Day", "Start Time", "Stop Time"]

    def __init__(self):
        forecast_headers = ["Event #", "Student Name", "Grade", "Event", "Time", "Day",
                            "Student Count", "Student:Instructor", "Instructors Required"]
        schedule_headers = ["Instructor Name", "Day", "Start Time", "Stop Time"]

    def write_all(cls, events, instructors, forecast_detailed_ws, forecast_summary_ws, schedule_by_name_ws):
        cls.write_summary_forecast(events, forecast_summary_ws)
        cls.write_detailed_forecast(events, forecast_detailed_ws)
        cls.write_by_name_schedule(instructors, schedule_by_name_ws)

    def write_detailed_forecast(self, events, forecast_detailed_ws):
        print("\n Writing Schedule, Summary Forecast, Detailed Forecast")
        forecast_headers = ["Event #", "Student Name", "Grade", "Event", "Time", "Day",
                            "Student Count", "Student:Instructor", "Instructors Required"]
        row_num = 1
        col_num = 1
        for col_header in forecast_headers:
            forecast_detailed_ws.cell(row_num, col_num).value = col_header
            col_num = col_num + 1
        for each_event in events:
            row_num = row_num + 1
            col_num = 1
            for datum in each_event.tuple():
                forecast_detailed_ws.cell(row_num, col_num).value = datum
                col_num = col_num + 1

    def write_summary_forecast(self, events, forecast_summary_ws):
        # Write Summary Forecast
        row_num = 1
        col_num = 1
        for col_header in self.forecast_headers:
            forecast_summary_ws.cell(row_num, col_num).value = col_header
            col_num = col_num + 1
        for each_event in events:
            if each_event.is_summary_event():
                row_num = row_num + 1
                col_num = 1
                for datum in each_event.tuple():
                    forecast_summary_ws.cell(row_num, col_num).value = datum
                    col_num = col_num + 1

    def write_by_name_schedule(self, instructors, schedule_by_name_ws):
        row_num = 1
        col_num = 1
        for col_header in self.schedule_headers:
            schedule_by_name_ws.cell(row_num, col_num).value = col_header
            col_num = col_num + 1
        for each_instructor in instructors:
            for each_day in each_instructor.schedule.keys():
                if each_instructor.schedule[each_day]:
                    row_num = row_num + 1
                    col_num = 1
                    for datum in each_instructor.tuple(each_day):
                        schedule_by_name_ws.cell(row_num, col_num).value = datum
                        col_num = col_num + 1

    def write_all(self, events, instructors, forecast_detailed_ws, forecast_summary_ws, schedule_by_name_ws):
        self.write_summary_forecast(events, forecast_summary_ws)
        self.write_detailed_forecast(events, forecast_detailed_ws)
        self.write_by_name_schedule(instructors, schedule_by_name_ws)
