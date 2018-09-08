from tkinter import Tk, filedialog, simpledialog


class Common:

    # define days consistent with datetime.weekday()
    def __init__(self):
        mon = 0
        tue = 1
        wed = 2
        thu = 3
        fri = 4
        sat = 5
        sun = 6

        root = Tk()
        root.withdraw()

        # string lookup for days of the week
        dayStrings = {sun: "Sunday", mon: "Monday", tue: "Tuesday", wed: "Wednesday", thu: "Thursday", fri: "Friday", \
                      sat: "Saturday"}


