from classes import *

window = GUI()

button_list_student = Button(text="SELECT STUDENT LIST", command=window.student_list)
window.config_button_xls(button_list_student)
button_list_student.place(x=5, y=50)

button_reports = Button(text="SELECT REPORTS FOLDER", command=window.reports)
window.config_button_folder(button_reports)
button_reports.place(x=5, y=150)

button_answers = Button(text="SELECT ANSWERS KEY FOLDER", command=window.answers)
window.config_button_folder(button_answers)
button_answers.place(x=5, y=250)

button_start = Button(text="START", command=window.start_process)
window.config_button_folder(button_start)
button_start.place(x=5, y=350)

button_display_result = Button(text="DISPLAY RESULT", command=window.display_result_file)
window.config_button_folder(button_display_result)
button_display_result.place(x=5, y=450)


window.compile_window()