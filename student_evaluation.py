import tkinter as tk
from tkinter import ttk
from tkinter import Menu
from tkinter import messagebox
from tkinter import filedialog
from tkinter import font
import json
from docx import Document
from docx.shared import Inches,RGBColor,Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.dml.color import ColorFormat

categories = ["Content structure / ideas", "Language and Delivery",  "Technical"]
colours = {
    4: "#77FF77",
    3: "#FFFF77",
    2: "#FFAA77",
    1: "#FF6666",
    0: "#77CCCC",
}

evaluation_parameters = [
    [
        {
            "label": "Focus",
            "evaluations": [
                {
                    "label": "Excellent",
                    "value": 4,
                    "description": "Purpose of presentation is clear from the outset. Supporting ideas maintain clear focus on the topic.",
                },
                {
                    "label": "Good",
                    "value": 3,
                    "description": "Topic of the presentation is clear. Content generally supports the purpose",
                },
                {
                    "label": "Fair",
                    "value": 2,
                    "description": "Presentation lacks clear direction. Big ideas not specifically identified.",
                },
                {
                    "label": "Poor",
                    "value": 1,
                    "description": "No focus at all. Audience cannot determine purpose of presentation.",
                },
            ],
        },
        {
            "label": "Organization",
            "evaluations": [
                {
                    "label": "Excellent",
                    "value": 4,
                    "description": "Student presents information in logical, interesting sequence that audience follows",
                },
                {
                    "label": "Good",
                    "value": 3,
                    "description": "Student presents information in logical sequence that audience can follow",
                },
                {
                    "label": "Fair",
                    "value": 2,
                    "description": "Audience has difficulty following because student jumps around.",
                },
                {
                    "label": "Poor",
                    "value": 1,
                    "description": "Audience cannot understand because there is no sequence of information",
                },
            ],
        },
        {
            "label": "Visual Aids",
            "evaluations": [
                {
                    "label": "Excellent",
                    "value": 4,
                    "description": "Visual aids are readable, clear and professional looking, enhancing the message.",
                },
                {
                    "label": "Good",
                    "value": 3,
                    "description": "Visual aids are mostly readable, clear and professional looking.",
                },
                {
                    "label": "Fair",
                    "value": 2,
                    "description": "Significant problems with readability, clarity, professionalism of visual aids.",
                },
                {
                    "label": "Poor",
                    "value": 1,
                    "description": "Visual aids are all unreadable, unclear and/or unprofessional.",
                },
            ],
        },
        {
            "label": "Question & Answer",
            "evaluations": [
                {
                    "label": "Excellent",
                    "value": 4,
                    "description": "Speaker has prepared relevant questions for opening up the discussion and is able to stimulate discussion.",
                },
                {
                    "label": "Good",
                    "value": 3,
                    "description": "Speaker has prepared relevant questions for opening up the discussion and is somewhat able to stimulate discussion",
                },
                {
                    "label": "Fair",
                    "value": 2,
                    "description": "Speaker has prepared questions but is not really able to stimulate discussion",
                },
                {
                    "label": "Poor",
                    "value": 1,
                    "description": "Speaker has not prepared questions",
                },
            ],
        },
    ],
    [
        {
            "label": "Eye Contact",
            "evaluations": [
                {
                    "label": "Excellent",
                    "value": 4,
                    "description": "Holds attention of entire audience with the use of direct eye contact, seldom looking at notes",
                },
                {
                    "label": "Good",
                    "value": 3,
                    "description": "Consistent use of direct eye contact with audience, but often returns to notes",
                },
                {
                    "label": "Fair",
                    "value": 2,
                    "description": "Displays minimal eye contact with audience, while reading mostly from the notes.",
                },
                {
                    "label": "Poor",
                    "value": 1,
                    "description": "No eye contact with audience; entire presentation is read from notes",
                },
            ],
        },
        {
            "label": "Enthusiasm",
            "evaluations": [
                {
                    "label": "Excellent",
                    "value": 4,
                    "description": "Demonstrates a strong, positive feeling about topic during entire presentation.",
                },
                {
                    "label": "Good",
                    "value": 3,
                    "description": "Mostly shows positive feelings about topic.",
                },
                {
                    "label": "Fair",
                    "value": 2,
                    "description": "Shows some negativity toward topic presented.",
                },
                {
                    "label": "Poor",
                    "value": 1,
                    "description": "Shows no interest in topic presented.",
                },
            ],
        },
        {
            "label": "Elocution",
            "evaluations": [
                {
                    "label": "Excellent",
                    "value": 4,
                    "description": "Student uses a clear voice so that all audience members can hear presentation",
                },
                {
                    "label": "Good",
                    "value": 3,
                    "description": "Student’s voice is clear. Most audience members can hear presentation.",
                },
                {
                    "label": "Fair",
                    "value": 2,
                    "description": "Student’s voice is low. Audience has difficulty hearing presentation.",
                },
                {
                    "label": "Poor",
                    "value": 1,
                    "description": "Student mumbles, speaks too quietly for a majority of audience to hear.",
                },
            ],
        },
    ],
    [
        {
            "label": "Knowledge",
            "evaluations": [
                {
                    "label": "Excellent",
                    "value": 4,
                    "description": "Demonstrate clear knowledge and understanding of the subject",
                },
                {
                    "label": "Good",
                    "value": 3,
                    "description": "Show clear knowledge and understanding of most subject area",
                },
                {
                    "label": "Fair",
                    "value": 2,
                    "description": "Show some knowledge and understanding of the subject area",
                },
                {
                    "label": "Poor",
                    "value": 1,
                    "description": "Show no knowledge and understanding of the subject area",
                },
            ],
        },
        {
            "label": "Research",
            "evaluations": [
                {
                    "label": "Excellent",
                    "value": 4,
                    "description": "Evidence of thorough research and preparation",
                },
                {
                    "label": "Good",
                    "value": 3,
                    "description": "Evidence of sufficient research and preparation",
                },
                {
                    "label": "Fair",
                    "value": 2,
                    "description": "Evidence of some research and preparation",
                },
                {
                    "label": "Poor",
                    "value": 1,
                    "description": "Evidence of no research and preparation",
                },
            ],
        },
        {
            "label": "Discussion of new ideas",
            "evaluations": [
                {
                    "label": "Excellent",
                    "value": 4,
                    "description": "Demonstrate  thorough knowledge while discussing new ideas",
                },
                {
                    "label": "Good",
                    "value": 3,
                    "description": "Show  sufficient knowledge while discussing new ideas",
                },
                {
                    "label": "Fair",
                    "value": 2,
                    "description": "Show  some knowledge while discussing new ideas",
                },
                {
                    "label": "Poor",
                    "value": 1,
                    "description": "Show  no knowledge while discussing new ideas",
                },
            ],
        },
        {
            "label": "Argument",
            "evaluations": [
                {
                    "label": "Excellent",
                    "value": 4,
                    "description": "Opinion set out in a concise and persuasive manner",
                },
                {
                    "label": "Good",
                    "value": 3,
                    "description": "Opinion is not concise and persuasive manner",
                },
                {
                    "label": "Fair",
                    "value": 2,
                    "description": "Opinion is clearly demonstrated but not persuasive ",
                },
                {
                    "label": "Poor",
                    "value": 1,
                    "description": "Opinion is not demonstrated or highlighted",
                },
            ],
        },
        {
            "label": "Questions",
            "evaluations": [
                {
                    "label": "Excellent",
                    "value": 4,
                    "description": "Responded very well to technical questions",
                },
                {
                    "label": "Good",
                    "value": 3,
                    "description": "Could answer most technical questions related to the presentation",
                },
                {
                    "label": "Fair",
                    "value": 2,
                    "description": "Could answer some technical questions related to the presentation",
                },
                {
                    "label": "Poor",
                    "value": 1,
                    "description": "Could not answer any technical questions related to the presentation",
                },
            ],
        },
    ],
]

evaluate_radios = []
evaluate_descriptions = []
evaluate_values = []
total_marks = 0

def get_result_string(result_marks):
    format = "Result: {result_marks:d} / {total_marks}"
    mark_string = format.format(result_marks = result_marks, total_marks = total_marks)
    #print(mark_string)
    return mark_string

def get_marks():
    sum = 0
    y = 0
    for i in range(len(categories)):
        for j in range(len(evaluation_parameters[i])):
            evaluation_parameter = evaluation_parameters[i][j]
            evaluate_count = len(evaluation_parameter["evaluations"])

            selected = evaluate_values[y].get()
            if (selected < evaluate_count):
                sum = sum + evaluation_parameter["evaluations"][selected]["value"]

            y = y + 1

    return sum

def click_evaluation(y, k):
    def x():
        result_label["text"] = get_result_string(get_marks())
        #messagebox.showinfo("title", "body" + str(y) + str(k))
    return x

def click_evaluation_selection(clicked_i, clicked_j, clicked_y, evaluate_count):
    def x():
        y = 0
        for i in range(len(categories)):
            for j in range(len(evaluation_parameters[i])):
                evaluation_parameter = evaluation_parameters[i][j]
                evaluate_count = len(evaluation_parameter["evaluations"])
                for k in range(len(evaluation_parameter["evaluations"]) + 1):
                    if y == clicked_y:
                        #evaluate_radios[y][k].configure(state = tk.NORMAL)
                        evaluate_descriptions[y][k].grid()
                    #else:
                        #evaluate_radios[y][k].configure(state = tk.DISABLED)
                        #evaluate_descriptions[y][k].grid_remove()
                y = y + 1
    return x

def create_result_json_dict():
    selected = []
    for i in range(len(evaluate_values)):
        selected.append(evaluate_values[i].get())
    
    data = {
        "student" : {
            "name" : student_name_first_entered.get() + " " + student_name_last_entered.get(),
            "id" : student_id_entered.get(),
        },
        "topic" : topic_entered.get(),
        "total_marks" : total_marks,
        "selects" : selected,
    }

    return data

def create_result_json():
    data = create_result_json_dict()
    return json.dumps(data)

def validate_fill_in():
    message = ""
    result = True

    if student_name_first_entered.get() == "":
        message += "* Missing Student name (first name)\n"

    if student_name_last_entered.get() == "":
        message += "* Missing Student name (last name)\n"

    if student_id_entered.get() == "":
        message += "* Missing Student ID\n"

    if topic_entered.get() == "":
        message += "* Missing Topic of presentation\n"

    y = 0
    for i in range(len(categories)):
        for j in range(len(evaluation_parameters[i])):
            evaluation_parameter = evaluation_parameters[i][j]
            evaluate_count = len(evaluation_parameter["evaluations"])

            if evaluate_count < evaluate_values[y].get():
                txt = "* {label:s} is not assessed.\n"
                message += txt.format(label = evaluation_parameter["label"])

            y = y + 1


    if (message != ""):
        messagebox.showinfo("Error", message)
        result = False

    return result

def menu_save():
    if (validate_fill_in()):
        json = create_result_json()

        txt = "{student_id:s}.json"
        file_name = txt.format(txt, student_id = student_id_entered.get())

        with open(file_name, mode='w') as f:
            f.write(json)

def menu_export_as_word():
    if (validate_fill_in()):
        txt = "{student_id:s}.docx"
        filename = txt.format(student_id = student_id_entered.get())
        path = tk.filedialog.asksaveasfile(mode = "w", initialfile = filename, defaultextension = "docx")
        filename = path.name
        path.close()
        do_exportAsWord(filename)

def do_exportAsWord(filename):
    j = create_result_json_dict()

    studentId = j["student"]["id"]

    txt = "{result_marks:d} / {total_marks:d}"
    mark = txt.format(result_marks = get_marks(), total_marks = total_marks)
    doc = Document()
    doc.add_paragraph('Student name: ' + j["student"]["name"] + "\t" + 'Student ID: ' + j["student"]["id"] + "\t\tResult: " + mark)
    doc.add_paragraph('Topic of Presentation: ' + j["topic"])

    table = doc.add_table(rows = 15, cols = 5)
    table.style = 'TableGrid'

    rowIndex = 0
    jIndex = 0

    for category_idx in range(len(categories)):
        table.cell(rowIndex, category_idx).text = categories[category_idx]
        table.cell(rowIndex, category_idx).paragraphs[0].runs[0].font.bold = True

        #
        # label
        #
        for category_row_index in range(len(evaluation_parameters[category_idx])):
            options = evaluation_parameters[category_idx][category_row_index]["evaluations"]

            for option_index in range(len(options)):
                option = options[option_index]
                #print(option)
                txt = "{score:d} - {label:s}"
                table.cell(rowIndex, option_index + 1).text = txt.format(score = option["value"], label = option["label"])
                table.cell(rowIndex, option_index + 1).paragraphs[0].runs[0].font.bold = True

        for category_row_index in range(len(evaluation_parameters[category_idx])):
            options = evaluation_parameters[category_idx][category_row_index]["evaluations"]
            rowIndex = rowIndex + 1
            option = evaluation_parameters[category_idx][category_row_index]
            table.cell(rowIndex,0).text = option["label"]
            selected = j["selects"][jIndex]

            for option_index in range(len(options)):
                option = options[option_index]
                table.cell(rowIndex, option_index + 1).text = option["description"]
                if (selected == option_index):
                    table.cell(rowIndex, option_index + 1).paragraphs[0].runs[0].font.bold = True

            jIndex = jIndex + 1


    doc.save(filename)
    messagebox.showinfo("title", "save as " + filename)

def menu_exit():
    win.quit()
    win.destroy()
    exit()

def menu_about():
    messagebox.showinfo("About", "This application is for student evaluation created by SUX2 group in SOFT808 in S2 2020.")

# Create instance
win = tk.Tk()
win.resizable(0, 0)

# Add a title
win.title("Student Evaluation")

default_font = tk.font.nametofont("TkDefaultFont")
small_font = tk.font.Font(size = default_font.cget("size") - 1)

#
#  Frames
#

student_frame = ttk.LabelFrame(win, text = "Student")
student_frame.grid(column = 0, row = 0, padx = 8, pady = 4, sticky = tk.W)
#student_frame.pack(expand = True, fill = tk.X)

result_frame = ttk.LabelFrame(win, text = "Result")
result_frame.grid(column = 1, row = 0, padx = 8, pady = 4, sticky = tk.E + tk.W + tk.N + tk.S)
#result_frame.pack(side = tk.RIGHT, expand = True, fill = tk.X)

evaluation_frame = ttk.LabelFrame(win, text = "Evaluation")
evaluation_frame.grid(column = 0, row = 1, columnspan = 2, padx = 8, pady = 4)
#evaluation_frame.pack(side = tk.BOTTOM, expand = True, fill = tk.X)

#
#  Student name
#
label = ttk.Label(student_frame, text = "Student name", )
label.grid(column = 0, row = 1, sticky = tk.W, ipady = 4)

label = ttk.Label(student_frame, text = "First name", font = small_font)
label.grid(column = 1, row = 0, sticky = tk.W)
label = ttk.Label(student_frame, text = "Last name", font = small_font)
label.grid(column = 2, row = 0, sticky = tk.W)

student_first_name = tk.StringVar()
student_last_name = tk.StringVar()
student_name_first_entered = ttk.Entry(student_frame, width = 15, textvariable = student_first_name)
student_name_first_entered.grid(column = 1, row = 1, sticky = tk.W)
student_name_last_entered = ttk.Entry(student_frame, width = 15, textvariable = student_last_name)
student_name_last_entered.grid(column = 2, row = 1, sticky = tk.W)

#
#  Student ID
#
label = ttk.Label(student_frame, text = "Student ID")
label.grid(column = 5, row = 1, sticky = tk.W, ipadx = 10, ipady = 4)

student_id = tk.StringVar()
student_id_entered = ttk.Entry(student_frame, width = 12, textvariable = student_id)
student_id_entered.grid(column = 6, row = 1, sticky = tk.W)

#
#  Topic of presentation
#
label = ttk.Label(student_frame, text = "Topic of presentation")
label.grid(column = 0, row = 2, sticky = tk.W, ipady = 4)

topic = tk.StringVar()
topic_entered = ttk.Entry(student_frame, width = 30, textvariable = topic)
topic_entered.grid(column = 1, row = 2, columnspan = 2, sticky = tk.W)

#
#  Result
#
result_label = ttk.Label(result_frame, text = get_result_string(0), width = 30)
result_label.grid(column = 0, row = 0, sticky = tk.E)

export_button = ttk.Button(result_frame, text = "Export as Word file", command = menu_export_as_word, width = 20)
export_button.grid(column = 1, row = 0, sticky = tk.E + tk.N + tk.S)

y = 0

tab_control = ttk.Notebook(evaluation_frame)
tab_control.grid(column = 0, row = 0)
tab_frames = []
for i in range(len(categories)):
    category_frame = ttk.LabelFrame(tab_control)
    tab_control.add(category_frame, text = categories[i])
    #category_frame.grid(column = 0, row = y * 2, sticky = tk.W)
    tab_frames.append(category_frame)

    for j in range(len(evaluation_parameters[i])):
        evaluate_values.append(tk.IntVar())
        evaluate_radios.append([])
        evaluate_descriptions.append([])
        evaluation_parameter = evaluation_parameters[i][j]
        evaluate_count = len(evaluation_parameter["evaluations"])
        evaluate_values[y].set(evaluate_count + 1)

        max_mark = 0
        for k in range(evaluate_count):
            colour = colours[evaluation_parameter["evaluations"][k]["value"]]
            label = str(evaluation_parameter["evaluations"][k]["value"]) + " - " + evaluation_parameter["evaluations"][k]["label"]
            r = tk.Radiobutton(category_frame, text = label, variable = evaluate_values[y], width = 20, bg = colour,
                                    value = k, command = click_evaluation(y, k))
            r.grid(column = k + 2, row = y * 2, sticky = tk.W)
            #r.configure(state = tk.DISABLED)

            description = evaluation_parameter["evaluations"][k]["description"]
            d = tk.Message(category_frame, text = description, bg = colour, font = small_font)
            d.grid(column = k + 2, row = y * 2 + 1, sticky = tk.W + tk.E + tk.N + tk.S)
            #d.grid_remove()

            evaluate_radios[y].append(r)
            evaluate_descriptions[y].append(d)

            if max_mark < evaluation_parameter["evaluations"][k]["value"]:
                max_mark = evaluation_parameter["evaluations"][k]["value"]

        total_marks = total_marks + max_mark

        label = "Not yet"
        print(evaluate_count)
        r = tk.Radiobutton(category_frame, text = label, variable = evaluate_values[y],
                                    value = evaluate_count + 1, bg = colours[0], command = click_evaluation(y, k))
        r.grid(column = evaluate_count + 2, row = y * 2, sticky = tk.W)
        #r.configure(state = tk.DISABLED)

        d = tk.Message(category_frame, text = "", bg = colours[0], font = small_font)
        d.grid(column = evaluate_count + 2, row = y * 2 + 1, sticky = tk.W + tk.E + tk.N + tk.S)
        #d.grid_remove()

        evaluate_radios[y].append(r)
        evaluate_descriptions[y].append(d)

        category_label = ttk.Label(category_frame, text = evaluation_parameter["label"], width = 20)
        category_label.grid(column = 1, row = y * 2, sticky = tk.W)

        y = y + 1


result_label["text"] = get_result_string(0)

menu_bar = Menu(win)
win.config(menu = menu_bar)

file_menu = Menu(menu_bar, tearoff = 0)
#file_menu.add_command(label = "Save", command = menu_save)
file_menu.add_command(label = "Export as Word", command = menu_export_as_word)
file_menu.add_separator()
file_menu.add_command(label = "Exit", command = menu_exit)
menu_bar.add_cascade(label = "File", menu = file_menu)

help_menu = Menu(menu_bar, tearoff = 0)
help_menu.add_command(label = "About", command = menu_about)
menu_bar.add_cascade(label = "Help", menu = help_menu)

win.mainloop()
