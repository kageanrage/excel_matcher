import sqlite3
import pandas as pd
import os
import csv
import xlsxwriter
import sys
import send2trash
import logging
logging.basicConfig(level=logging.DEBUG, format=' %(asctime)s - %(levelname)s - %(message)s')
# logging.disable(logging.CRITICAL)     # switches off logging

# Master TODO list
# TODO PENDING: open chrome, log in, go to admin page, desired project, download current results
# TODO DONE: convert that project csv into a project xlsx
# TODO DONE: convert that project xlsx into a project DB
# TODO DONE: connect to the SurveyResults database, attach All_Members DB, get a list of member IDs from SurveyResults DB
# TODO DONE: capitalise the member IDs in the DB
# TODO DONE: look up age and gender from Member table and add them to project table
# TODO DONE: convert project table to project csv
# TODO DONE: close DB connection
# TODO DONE: convert project csv to project xlsx
# TODO DONE: delete un-needed files


# key variable names
# in All_Members.db, table name is All_Members, ID column is ID, other col's of interest are ID, Age, Gender, Region, SES Member, email list?
# in SurveyResults.csv, tab name is SurveyResults, ID column is MemberId, other COI is AttemptId


# TODO: need to refactor to remove the hard-coding of filenames and other variables
fname = 'SR'


# this fn converts csv to xlsx
def csv_to_xlsx(input_filename, output_filename):
    wb = xlsxwriter.Workbook(output_filename + ".xlsx")
    ws = wb.add_worksheet(input_filename)    # desired worksheet title here
    with open(input_filename + ".csv", 'r') as csvfile:
        table = csv.reader(csvfile)
        i = 0
        # write each row from the csv file as text into the excel file
        # this may be adjusted to use 'excel types' explicitly (see xlsxwriter doc)
        for row in table:
            ws.write_row(i, 0, row)
            i += 1
    wb.close()


# this fn converts xls to SQLite DB
def xls_to_sql(filename):
    # print(f'attempting to open {filename + ".db"}')
    con = sqlite3.connect(filename + ".db")
    wb = pd.read_excel(filename + '.xlsx', sheet_name=None)
    for sheet in wb:
       wb[sheet].to_sql(sheet, con, index=False)
    con.commit()
    con.close()


def delete_files(*args):
    for arg in args:
        send2trash.send2trash(arg)
        logging.debug(f'{arg} sent to trash')






os.chdir("H:\WorkingDir\member_data")  # set CWD
direc = os.getcwd()
logging.debug(f'Setting CWD to {direc}')




# TODO DONE: convert that project csv into a project xlsx
# now converting csv to xlsx
logging.debug(f'now converting csv to xlsx')
# csv_to_xlsx("SurveyResults", "SurveyResults") # converts csv to xlsx, parameters are desired input and output filenames
csv_to_xlsx(fname, fname)  # converts csv to xlsx, parameters are desired input and output filenames

# TODO DONE: convert that project xlsx into a project DB

# xls_to_sql("All_Members")  # convert All Members xls to db
logging.debug(f'now converting xlsx to sql')
# xls_to_sql("SurveyResults")  # convert SurveyResults xls to db
xls_to_sql(fname)  # convert SurveyResults xls to db

# TODO DONE: connect to the SurveyResults database, attach All_Members DB, get a list of member IDs from SurveyResults DB
# This is my main initial section, connecting the databases.

db_file = fname + ".db"

logging.debug(f'initiating DB connection and cursor')
conn = sqlite3.connect(db_file)
c = conn.cursor()

logging.debug(f'Attaching All_Members database')
c.execute("ATTACH DATABASE 'All_Members.db' AS All_Members")  # attaches SurveyResults db to the connection, which I later access via read_from_db2()

logging.debug(f'Adding columns to {fname} db for Age and Gender')
c.execute("ALTER TABLE {tn} ADD COLUMN '{cn}' {ct}".format(tn=fname, cn='Age', ct='INTEGER'))  # Add (empty) Age column to SurveyResults db
c.execute("ALTER TABLE {tn} ADD COLUMN '{cn}' {ct}".format(tn=fname, cn='Gender', ct='TEXT'))  # Add (empty) Gender column to SurveyResults db

logging.debug(f'Selecting all member IDs from {fname} DB')
c.execute("SELECT {coi} from {tn}".format(tn=fname, coi='MemberId'))  # generates a list of member IDs from db
all_member_IDs = c.fetchall()
members_list_length = len(all_member_IDs)
assert members_list_length > 0, 'Length of members_list is not > 0'
logging.debug(f'member_IDs list length is {members_list_length}')
# logging.debug('member IDs are as follows:')
# pprint.pprint(all_member_IDs)


# TODO DONE: capitalise the member IDs in the DB

# This section of code goes through SurveyResults and overwrites current member IDs with upper case version
# for id_number in test_member_IDs:
logging.debug(f'Capitalising Member IDs in {fname} DB')
for id_number in all_member_IDs:
    extracted_id = id_number[0]
    logging.debug(f'Extracted ID is {extracted_id}')
    # id_upper = id_number.upper()
    id_upper = id_number[0].upper()
    logging.debug(f'id_upper is {id_upper}')
    c.execute("UPDATE {tn} SET {cn}=('{data}') WHERE {lcn}=('{idno}')".format(tn=fname, cn='MemberId', data=id_upper, lcn='MemberId', idno=extracted_id))

conn.commit()


# TODO DONE: look up age and gender from Member table and add them to project table
# This section loops through IDs in SurveyResults, looks them up in All_Members and then updates age + gender in SurveyResults

# for id_number in test_member_IDs_5_to_10:  # for test mode or iterate through all_member_IDs which are each tiny tuples hence the need for [0] below
logging.debug('Looking up member IDs in All_Members to grab age and gender')
for id_number in all_member_IDs:
    extracted_id = id_number[0]  # grab the ID out of the weird little tuple and call it extracted_id
    extracted_id = extracted_id.upper()  # Converted to upper case so that the search works. Consider giving this a new variable name instead, like id_upper as in prev section
    # extracted_id = id_number    # or in test mode, no need to extract as it's just from the test list, but name it extracted id anyway
    c.execute("SELECT {coi1}, {coi2} FROM {tn} WHERE {lcn}='{idno}'".format(coi1='Age', coi2='Gender', tn='All_Members', lcn='Id', idno=extracted_id))  # look up ID in All_Members and select Age + gender
    try:
        age, gen = c.fetchone()  # assign age and gender to variables
    except TypeError:  # or, if the member ID wasn't found in All_Members
        print(f'Error - Member {extracted_id} not found')  # show an error message
        age = 0  # assign age to zero
        gen = 'Unknown'  # assign gender to unknown
    logging.debug(f'Checking ID {extracted_id}, age = {age}, gender = {gen}')  # just logging activity for troubleshooting
    c.execute("UPDATE {tn} SET {cn}=('{data}') WHERE {lcn}=('{idno}')".format(tn=fname, cn='Gender', data=gen, lcn='MemberId', idno=extracted_id))     # populates gender field in SurveyResults table
    c.execute("UPDATE {tn} SET {cn}=({data}) WHERE {lcn}=('{idno}')".format(tn=fname, cn='Age', data=age, lcn='MemberId', idno=extracted_id))  # populates age field in SurveyResults table
conn.commit()  # commit changes to DB


# TODO PENDING: convert project table to project xlsx
# now I'll borrow code to attempt to output db to excel
# this works but it converts to csv, not xlsx
logging.debug('Converting project table to project xlsx')

conn.row_factory = sqlite3.Row
# crsr = conn.execute("SELECT {coi1},{coi2},{coi3},{coi4},{coi5} FROM {tn}".format(coi1='MemberId', coi2='Outcome', coi3='FirstName', coi4='Age', coi5='Gender', tn=fname))  # if I want just select columns
crsr = conn.execute("SELECT * FROM {tn}".format(tn=fname))
row = crsr.fetchone()
titles = row.keys()

# data = c.execute("SELECT {coi1},{coi2},{coi3},{coi4},{coi5} FROM {tn}".format(coi1='MemberId', coi2='Outcome', coi3='FirstName', coi4='Age', coi5='Gender', tn=fname))  # if I want just select columns
data = c.execute("SELECT * FROM {tn}".format(tn=fname))
if sys.version_info < (3,):
    f = open(fname + 'exported.csv', 'wb')
else:
    f = open(fname + 'exported.csv', 'w', newline="")

writer = csv.writer(f, delimiter=',')
writer.writerow(titles)  # keys=title you're looking for
# write the rest
writer.writerows(data)
f.close()


# TODO DONE: close DB connection
# commit changes to SurveyResults database, close connection
logging.debug('Closing DB connection')
conn.commit()
c.close()



# TODO DONE: convert project csv to project xlsx
# now converting csv to xlsx, using this function for the 2nd time
logging.debug('now converting csv to xlsx, using this function for the 2nd time')
csv_to_xlsx(fname + "exported", fname + " appended")  # converts csv to xlsx, parameters are desired input and output filenames


# TODO DONE: delete un-needed files
delete_files(fname + 'exported.csv', fname + '.xlsx', fname + '.db')


