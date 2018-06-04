#!/usr/bin/python
import zipfile
import functions
import shutil
import os

zipped = zipfile.ZipFile(r"/var/www/html/uploads/compress.zip", 'r')
zipped.extractall(r"/var/www/html/uploads/compress")
os.system("chown www-data /var/www/html/uploads/compress")
newfinal = r"/var/www/html/cgi-bin/downloads/newcombined.xls"
path = r"/var/www/html/uploads/compress"
execute = r"/var/www/html/uploads/Execute"

for filenameList in os.listdir(path):
    if os.path.isdir(os.path.join(path, filenameList)):
        path = os.path.join(path, filenameList)
for x in os.listdir(path):
    os.chmod(path + '/' + x, 0o777)

for newfiles in os.listdir(path):
    if newfiles in os.listdir(execute):
        pass
    else:
        shutil.copy(path + '/' + newfiles, execute)

path = execute


def main_Funtion():
    print("Content-type:text/html\r\n\r\n")

    page = r"/var/www/html/downloadpage.html"
    file = open(page, 'r')
    print(file.read())
    functions.combing_files(path, newfinal)
    functions.top_row(newfinal)
    functions.drop_duplicates(newfinal, "last")
    functions.end_time(newfinal)

    functions.new_resolved_cols(newfinal)
    functions.res_row(newfinal)

    functions.correct_dept(newfinal)
    functions.end_state(newfinal)
    functions.format_date(newfinal)
    functions.drop_duplicates(newfinal, "last")
    functions.time_diff(path, newfinal)
    functions.aggregate(newfinal)
    functions.timetaken(newfinal)
    functions.database_operations()
    # functions.testing(path, newfinal)

    shutil.rmtree(r"/var/www/html/uploads/compress")
    print('</body></html>')
    return


main_Funtion()
