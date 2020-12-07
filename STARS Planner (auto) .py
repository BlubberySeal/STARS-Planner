import PySimpleGUI as sg
import itertools
from tkinter import *
from tkinter import ttk
from selenium import webdriver


# Here's the basic idea of how this program works:
# 1. Extracts all the data from the Excel file OR web-scrapes the horrendously slow STARS website.
# 2. Generates ALL possible timetables. Any timetable with a period clash is automatically discarded.
# 3. The master timetable list is filtered according to user's selection. Master timetable list is NEVER re-generated.
# 4. Displays the timetables with the power of magic (and quite possibly Tkinter).


# Converts the time range (eg. 1430-1630) to a grid format that can be interpreted by the program
def process_time(time):
    start_time = int(time[:4])
    end_time = int(time[5:])

    if start_time % 50 != 0:
        start_time += 20

    if end_time % 50 != 0:
        end_time += 20

    start_row = (start_time - 800) / 50 + 1
    end_row = (start_row + (end_time - start_time) / 50)

    return [start_row, end_row]


# Converts the day (eg. "Mon") to a grid format that can be interpreted by the program
def process_day(which_day):
    for index, name in enumerate(["Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]):
        if which_day == name:
            return index + 1


# I'm too lazy to write 6 different labels for each day so this function is here
def give_day(number):
    for index, name in enumerate(["Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]):
        if number == index:
            return name


# Still too lazy to write 31 different labels for each time period
def give_time(number):
    if number % 2 == 0:
        product = number * 50
        text = "{:04d}-{:04d}".format(product + 800, product + 830)
    else:
        product = (number-1) * 50 + 30
        text = "{:04d}-{:04d}".format(product + 800, product + 870)

    return text


# Increases the timetable no. to display by 1
def forward(table_number):
    global timetable_no

    table_number += 1
    timetable_no = table_number
    put_base()


# Decreases the timetable no. to display by 1
def back(table_number):
    global timetable_no

    table_number -= 1
    timetable_no = table_number
    put_base()


# Jumps to specified timetable no.
def goto(entry_field):
    global timetable_no
    global filtered_timetable_list

    table_to_go = int(entry_field.get())

    # Move to timetable no. only if input is valid
    if table_to_go in range(1, len(filtered_timetable_list)+1):
        timetable_no = table_to_go
        put_base()


# Filters the master timetable list based on selected criteria
def filter_tables():
    global list_of_timetables
    global filtered_timetable_list
    global timetable_no
    global period_var
    global day_var

    # Retrieves checkbox on/off state
    period_values = [i.get() for i in period_var]
    day_values = [i.get() for i in day_var]

    # Holds the filtered timetables
    temporary_timetable_list = []

    # I had to use "timetable_get" because "timetable" is another variable name in the global scope
    # I know, its lazy but I don't care :)
    for timetable_get in list_of_timetables:
        marker = 0

        # day_values[6] is the on/off variable toggled by the "Online" checkbox
        if day_values[6] != 1:
            for i in range(6):

                # There are 6 days, i = 0 -> Monday, i = 1 -> Tuesday, etc.
                if day_values[i] == 1:

                    # "row" is some variable from the global scope as well, hence "row_get"
                    for row_get in timetable_get:

                        # Check if there are lessons on the day
                        # i+1 represents the column/day
                        if row_get[i + 1][0] != 0:
                            marker = 1

            # One for each time period
            for i in range(31):

                # There are a total of 31 time periods, starting from 0800-0830
                if period_values[i] == 1:

                    # timetable_get[i] represents the row/time
                    for column in timetable_get[i + 1]:

                        # Check if there are lessons during that time period
                        if column[0] != 0:
                            marker = 1

        # "Online" is checked: any online lesson is treated as invisible
        elif day_values[6] == 1:
            for i in range(6):
                if day_values[i] == 1:

                    for row_get in timetable_get:
                        # Check if lessons exist on that day
                        if row_get[i + 1][0] != 0 and row_get[i + 1][3] != "ONLINE":
                            marker = 1

            for i in range(31):
                if period_values[i] == 1:

                    # timetable_get[i] represents the row/time
                    for column in timetable_get[i + 1]:
                        if column[0] != 0 and column[3] != "ONLINE":
                            marker = 1

        # If the timetable isn't filtered out, add it to filtered_timetable_list
        if marker == 0:
            temporary_timetable_list.append(timetable_get)

    # filtered_timetable_list are the timetables we will display
    filtered_timetable_list = temporary_timetable_list.copy()

    # Reset the timetable no. so it doesn't go out of bounds
    timetable_no = 1


def put_base():
    global filtered_timetable_list
    global timetable_no
    global period_var
    global day_var

    # If you want to see the widgets stack up and lag the shit of out your PC remove the following two lines :)
    for widget in root.winfo_children():
        widget.destroy()

    mainframe = Frame(root)
    mainframe.pack(fill=BOTH, expand=1)

    my_canvas = Canvas(mainframe)
    my_canvas.pack(side=LEFT, fill=BOTH, expand=1)

    scrollbar = ttk.Scrollbar(mainframe, orient=VERTICAL, command=my_canvas.yview)
    scrollbar.pack(side=RIGHT, fill=Y)
    my_canvas.config(yscrollcommand=scrollbar.set)
    my_canvas.bind('<Configure>', lambda e: my_canvas.configure(scrollregion=my_canvas.bbox("all")))
    base = Frame(my_canvas)
    my_canvas.create_window((0, 0), window=base, anchor="nw")

    for i in zip(range(7), ["Time/Day", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]):
        Label(base, text=i[1], padx=10).grid(column=i[0], row=0)

    for i in range(31):
        Label(base, text=give_time(i)).grid(column=0, row=i + 1)

    button_back = Button(base, text="<", command=lambda: back(timetable_no))
    button_forward = Button(base, text=">", command=lambda: forward(timetable_no))
    status = Label(base, text="Timetable {} of {}".format(timetable_no, len(filtered_timetable_list)), bd=2, relief=GROOVE, padx=10, pady=8)
    add10 = Button(base, text=">>", command=lambda: forward(timetable_no + 9))
    add100 = Button(base, text=">>>", command=lambda: forward(timetable_no + 99))
    back10 = Button(base, text="<<", command=lambda: back(timetable_no - 9))
    back100 = Button(base, text="<<<", command=lambda: back(timetable_no - 99))
    break_label = Label(base, text="------------------------")
    skip_label = Label(base, text="Go to timetable no.:")
    skip_entry = Entry(base)
    skip_button = Button(base, text="Go!", width=10, command=lambda: goto(skip_entry))
    break_label1 = Label(base, text="------------------------")
    check_label = Label(base, text="Keep which days free?")
    break_label2 = Label(base, text="------------------------")
    online_label = Label(base, text="ONLINE = Free Period")
    online_check = Checkbutton(base, text="Yes :)", variable=day_var[6])
    timetable_generate = Button(base, text="Generate!", command=lambda: [filter_tables(), put_base()])
    break_label3 = Label(base, text="------------------------")
    quit_button = Button(base, text="Quit Program", command=lambda: root.destroy())

    for i in range(31):
        btn = Checkbutton(base, variable=period_var[i], padx=15)
        btn.grid(column=7, row=i+1)

    for i in range(6):
        btn = Checkbutton(base, text=give_day(i), variable=day_var[i])

        if i % 2 == 0:
            btn.grid(column=8, row=int(11+(i/2)))

        else:
            btn.grid(column=9, row=int(10+(i/2 + 0.5)))

    if timetable_no == 1:
        button_back = Button(base, text="<", command=lambda: back(timetable_no), state=DISABLED)

    if timetable_no == len(filtered_timetable_list):
        button_forward = Button(base, text=">", command=lambda: forward(timetable_no), state=DISABLED)

    if timetable_no <= 10:
        back10 = Button(base, text="<<", command=lambda: back(timetable_no - 9), state=DISABLED)

    if timetable_no <= 100:
        back100 = Button(base, text="<<<", command=lambda: back(timetable_no - 99), state=DISABLED)

    if len(filtered_timetable_list) - timetable_no < 10:
        add10 = Button(base, text=">>", command=lambda: forward(timetable_no + 9), state=DISABLED)

    if len(filtered_timetable_list) - timetable_no < 100:
        add100 = Button(base, text=">>>", command=lambda: forward(timetable_no + 99), state=DISABLED)

    back100.grid(column=8, row=1)
    add100.grid(column=9, row=1)
    back10.grid(column=8, row=2)
    add10.grid(column=9, row=2)
    button_back.grid(column=8, row=3)
    button_forward.grid(column=9, row=3)
    status.grid(column=8, row=4, columnspan=2)
    break_label.grid(column=8, row=5, columnspan=2)
    skip_label.grid(column=8, row=6, columnspan=2)
    skip_entry.grid(column=8, row=7, columnspan=2)
    skip_button.grid(column=8, row=8, columnspan=2)
    break_label1.grid(column=8, row=9, columnspan=2)
    check_label.grid(column=8, row=10, columnspan=2)
    break_label2.grid(column=8, row=14, columnspan=2)
    online_label.grid(column=8, row=15, columnspan=2)
    online_check.grid(column=8, row=16, columnspan=2)
    timetable_generate.grid(column=8, row=17, columnspan=2)
    break_label3.grid(column=8, row=18, columnspan=2)
    quit_button.grid(column=8, row=19, columnspan=2)

    put_timetable(base)

    root.bind('<Return>', lambda event=None: goto(skip_entry))
    root.bind('<Left>', lambda event=None: back(timetable_no))
    root.bind('<Right>', lambda event=None: forward(timetable_no))


def put_timetable(frame):
    global filtered_timetable_list
    global timetable_no

    try:
        for i in range(32):
            for j in range(7):

                if i == 0 or j == 0:
                    pass
                # filtered_timetable_list[0] is the first timetable
                elif filtered_timetable_list[timetable_no - 1][i][j][0] != 0 and filtered_timetable_list[timetable_no - 1][i][j][3] == "ONLINE":

                    # Puts course code, index and venue into the cell
                    Label(frame, text="{}-{}\n{}".format(
                        filtered_timetable_list[timetable_no - 1][i][j][4],
                        filtered_timetable_list[timetable_no - 1][i][j][1],
                        filtered_timetable_list[timetable_no - 1][i][j][3]),
                        width=12, height=2, bd=1, fg="red", relief=GROOVE).grid(row=i, column=j)

                    # Puts course code, index and venue into the cell
                elif filtered_timetable_list[timetable_no - 1][i][j][0] != 0:
                    Label(frame, text="{}-{}\n{}".format(
                        filtered_timetable_list[timetable_no - 1][i][j][4],
                        filtered_timetable_list[timetable_no - 1][i][j][1],
                        filtered_timetable_list[timetable_no - 1][i][j][3]),
                        width=12, height=2, bd=1, relief=GROOVE).grid(row=i, column=j)

                else:
                    Label(frame, width=12, height=2, bd=1, relief=GROOVE).grid(row=i, column=j)

    # This is for when no timetables are filtered/generated (aka no combination possible)
    except IndexError:
        for i in range(32):
            for j in range(7):

                if i == 0 or j == 0:
                    pass

                else:
                    Label(frame, text="", width=12, height=2, bd=1, relief=GROOVE).grid(row=i, column=j)


main_dataset = {}
index_list = []
list_of_timetables = []
timetable_no = 1
filtered_timetable_list = []

sg.theme("SandyBeach")
layout = [[sg.Text("Username", size=(10, 1)), sg.InputText()],
          [sg.Text("Password", size=(10, 1)), sg.InputText(password_char='*')],
          [sg.Text("Courses", size=(10, 1)), sg.InputText()],
          [sg.OK()]]

window = sg.Window('Please enter info', layout)
user_input = window.read()[1]
window.close()

username = user_input[0]
password = user_input[1]
courses = user_input[2]

course_codes = courses.split()

driver = webdriver.Chrome()
driver.get('https://wish.wis.ntu.edu.sg/pls/webexe/ldap_login.login?w_url=https://wish.wis.ntu.edu.sg/pls/webexe/aus_stars_planner.main')

driver.find_element_by_id("UID").send_keys(username)
driver.find_element_by_id("UID").submit()
driver.find_element_by_id("PW").send_keys(password)
driver.find_element_by_id("PW").submit()
driver.find_element_by_xpath("//INPUT[@VALUE='Courses Selection and Info']").click()

for course in course_codes:

    table = []
    temp_index_list = []

    # Clear field, input course code and click "Search"
    driver.find_element_by_xpath("//INPUT[@Name='r_subj_code']").clear()
    driver.find_element_by_xpath("//INPUT[@Name='r_subj_code']").send_keys(course)
    driver.find_element_by_xpath("//INPUT[@VALUE='Search']").click()

    # Switch to iframe (the box containing the hyperlink)
    driver.switch_to.frame("subjects")

    # Find and click on the link (opens new tab)
    driver.find_element_by_partial_link_text(course).click()

    # Switch to new tab
    driver.switch_to.window(driver.window_handles[1])

    # Find the table with all index information
    table_xpath = driver.find_element_by_xpath("//table[@cellpadding='10']")

    # Iterate through table and stores info into a list
    for row in table_xpath.find_elements_by_xpath(".//tr"):

        row_list = []

        # Extract text row by row from the table
        for num in range(1, 7):
            row_list.append(row.find_elements_by_xpath(".//TD[{}]".format(num))[0].text.strip())

        # The table is represented using a 2D array
        table.append(row_list)

    # Removes the header row, we don't need that
    table.pop(0)

    # Data cleaning (some cells do not have the index)
    for row in range(len(table)):
        if table[row][0] == "":
            table[row][0] = table[row-1][0]

        # Populating index list (for generating combinations later)
        # eg. [ ['10000', '20000', '30000'], ['15000', '25000', '35000'] ...]
        #         ^ Indexes for MH1234         ^ Indexes for CZ9999
        if table[row][0] not in temp_index_list:
            temp_index_list.append(table[row][0])

            # Each index is a key to the dictionary to retrieve data for all its lessons
            main_dataset[table[row][0]] = []

        # One "temp_list" will contain information for ONE lesson
        # Each index may have multiple lessons
        # INDEX : [ [TYPE, DAY, [START, END], VENUE, COURSE], [TYPE, DAY, [START, END], VENUE, COURSE] ]
        table[row][4] = process_time(table[row][4])
        table[row][3] = process_day(table[row][3])
        table[row].append(course)

        # We don't need the group as well
        table[row].pop(2)
        main_dataset[table[row][0]].append(table[row][1:])

    index_list.append(temp_index_list)

    driver.close()
    driver.switch_to.window(driver.window_handles[0])

driver.quit()

index_combinations = list(itertools.product(*index_list))

for combination in index_combinations:

    early_break = False

    timetable = [[[0], [0, " ", " ", " ", " "], [0, " ", " ", " ", " "], [0, " ", " ", " ", " "], [0, " ", " ", " ", " "], [0, " ", " ", " ", " "], [0, " ", " ", " ", " "]] for i in range(32)]

    # Loop A
    for course_index in combination:

        # Loop B
        for lessons in main_dataset[course_index]:

            day = lessons[1]
            period_start = lessons[2][0]
            period_end = lessons[2][1]

            # Loop C
            for period in range(int(period_start), int(period_end)):

                timetable[period][day][0] += 1

                # Breaks out of Loop C if there is a clash
                if timetable[period][day][0] >= 2:
                    early_break = True
                    break

                # [OVERLAP, INDEX, TYPE, VENUE, COURSE]
                timetable[period][day][1] = course_index
                timetable[period][day][2] = lessons[0]
                timetable[period][day][3] = lessons[3]
                timetable[period][day][4] = lessons[4]

            # Breaks out of Loop B if there is a clash
            if early_break:
                break

        # Breaks out of Loop A if there is a clash
        # Tests next combination of indexes
        if early_break:
            break

    if not early_break:
        list_of_timetables.append(timetable)

filtered_timetable_list = list_of_timetables.copy()

root = Toplevel()
root.title("STARS Planner")
root.resizable(width=False, height=False)
root.geometry("820x1000")

day_var = [IntVar() for i in range(7)]
period_var = [IntVar() for i in range(31)]

put_base()

root.mainloop()
