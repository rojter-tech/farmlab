#!/usr/bin/python
import MySQLdb

db = MySQLdb.connect(host="rojter.tech",
                     user="dreuter",
                     passwd="**********",
                     db="dreuter")

cursor = db.cursor()

Query = """ LOAD DATA INFILE 'quandata.csv' \
INTO TABLE quandata_test \
FIELDS TERMINATED BY ';' \
LINES TERMINATED BY '\n' \
IGNORE 1 ROWS; """

cursor.execute(Query)
db.commit()
cursor.close()
