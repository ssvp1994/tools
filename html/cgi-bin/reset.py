#!/usr/bin/python
import mysql.connector
import os
import shutil
database = mysql.connector.connect(host='gsar1.karsun-csb.com', port='59100', user='csbops',
                                           password='csbopspswd', db="csbops")
cursor = database.cursor()
cursor.execute("delete from Fact_table;")

path=r"/var/www/html/uploads/Execute"
print("Content-type:text/html\r\n\r\n")
print("<html><body>")
cursor.close()
database.commit()
database.close()

for i in os.listdir(path):
    os.remove(path+'/'+i)
for x in os.listdir(r"/var/www/html/uploads"):
    if x=="compress":
        #os.system("chown www-data /var/www/html/uploads/compress")
        #os.system("chown www-data /var/www/html/uploads/compress/*")
        
        os.system("rm -rf /var/www/html/uploads/compress")

print("<h2><center> Database Reset Complete !!! </center></h2>")
print('</body></html>')