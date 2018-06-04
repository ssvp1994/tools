'''This program need only Excel files with .xls format, any other format won't work and specifically desined for weekly reports
of E.Tools - T2 (KARSUN SOLUTIONS)'''
import functions
import os



new= r"C:\Users\vgollapudi\Desktop"

final = r"\newcombined.xls"
newfinal = new+final

path =r"C:\Users\vgollapudi\Downloads\New folder\eTools Incident Metrics - 02062017-07242017"

for filenameList in os.listdir(path):
    if os.path.isdir(os.path.join(path,filenameList)):
        path=os.path.join(path,filenameList)


def main_Funtion():
    functions.combing_files(path,newfinal)

    functions.top_row(newfinal)
    functions.drop_duplicates(newfinal,"last")
    functions.end_time(newfinal)
    functions.new_resolved_cols(newfinal)
    functions.res_row(newfinal)
    functions.correct_dept(newfinal)
    functions.end_state(newfinal)
    functions.format_date(newfinal)
    functions.drop_duplicates(newfinal,"last")
    functions.time_diff(path,newfinal)
    functions.aggregate(newfinal)
    functions.timetaken(newfinal)
    functions.SLA_Counter(newfinal)
    #functions.testing(path,newfinal)



    return

main_Funtion()