#!/usr/bin/python
import MySQLdb 
import csv
import fnmatch
import os

db = MySQLdb.connect(host="rojter.tech",
                     user="dreuter",
                     passwd="**********",
                     db="farmlab")
cursor = db.cursor()

path = os.path.dirname(os.path.abspath(__file__)) + "\\"
print(path)
for file in os.listdir(path + '\csv'):
	if fnmatch.fnmatch(file,'*.csv'):
		print(file)
		Query = """ LOAD DATA LOCAL INFILE 'csv/%s' \
		INTO TABLE vbe_waters \
		FIELDS TERMINATED BY ';' \
		LINES TERMINATED BY '\n' \
		IGNORE 1 ROWS; """ % (file)
		cursor.execute(Query) 
		db.commit()
cursor.close()