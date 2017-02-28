'''
Created on Jan 15, 2013

@author: dongdong.zhang
'''
# Install Python for Windows Extension from http://sourceforge.net/projects/pywin32/
import win32com.client as win32
import os
import re
import sys
import getopt


# Load MS Word document and return the document object
def load_word_document(wordFile):
    word = win32.Dispatch("Word.Application")
    word.Visible = 0
    word.Documents.Open(wordFile)
    doc = word.ActiveDocument
    return doc


# The config file is plain text lines
# Each line has two fields separated by white spaces, the first field is index, the second field is the db table name
#
# Example:
# 3  cap_floor
# 4  coll_grp
# 5  coll_grp_delinq_hist
# 6  coll_grp_loan_hist
#
# Read each line and store it in a dictionary with the index as key and the table name as value.
def parse_config(configFile):
    file_object = open(configFile)
    all_lines = file_object.readlines()

    tableDict = {}
    for line in all_lines:
        pair = re.split('\s*', line)
        tableDict[pair[0]] = pair[1]

    file_object.close()

    return tableDict


def print_case_settings(file_object):
    file_object.write('*** Settings ***\n')
    file_object.write('Library  DatabaseLibrary\n\n')


def print_setup_section(file_object):
    file_object.write(
        "Suite Setup  Connect to Database Using Custom Params  Sybase  database = 'data_rigs', user = 'fis_read', passwd = 'fisreado', dsn ='${dsn}' \n\n")


def print_teardown_sectioin(file_object):
    file_object.write("Suite Teardown  Disconnect From Database\n\n")


def print_case_variables(file_object):
    file_object.write('*** Variables ***\n')
    file_object.write('${dsn}  EJV0STP1 \n\n')


def print_case_header(file_object):
    file_object.write('*** Test Cases ***\n')


def get_column_string(column):
    # convert to string
    columnStr = str(column)
    # Remove ^M('\r' on windows)
    columnStr = columnStr.replace('\r', '')
    # Remove ^G('\07' on windows)
    columnStr = columnStr.replace('\07', '')

    return columnStr


def print_case_body(file_object, tableName, table):
    print_case_settings(file_object)
    print_setup_section(file_object)
    print_teardown_sectioin(file_object)
    print_case_variables(file_object)
    print_case_header(file_object)

    numRows = table.Rows.Count

    allColumns = ""
    # Skip the first line since it is the table header
    for i in range(2, numRows + 1):
        # write the column name in a single line
        columnName = get_column_string(table.Cell(i, 1))
        allColumns += columnName + '  '

        # if the data type is not char or varchar, the length_of_column is set to 0
        # Get column type and length
        # varchar(10) = varchar  10
        # float       = float    0
        column_type = ''
        column_length = 0
        data = str(table.Cell(i, 3))
        if (data.find('(') != -1):
            pattern = re.compile('(\w+)\((\d*)\s*\)')
            match = pattern.match(data)
            column_type = match.group(1)
            column_length = match.group(2)
        else:
            column_type = data
            column_length = 0

        column_type = get_column_string(column_type)
        column_length = get_column_string(column_length)
        null_or_not = get_column_string(table.Cell(i, 4)).upper()

        # write the column check line, the format is
        # 'Column Check'  table_name  column_name  column_data_type  length_of_column
        # example:
        # Column Check  tranche  accrual_type_cd  varchar  10  NOT NULL
        columnContent = columnName + '\n' + '\tColumn Check' + '  ' + tableName + \
                        '  ' + columnName + '  ' + 'skip' + '  ' + column_length + '  ' + null_or_not + '\n'
        file_object.write(columnContent)

    # Print column match part
    columncheck = tableName + '\n' + '\tColumn Match' + '  ' + tableName + '  ' + allColumns
    file_object.write(columncheck)


def generate_Robot_test_cases(root, wordFile, configFile):
    doc = load_word_document(wordFile)
    tableConfig = parse_config(configFile)

    # Loop each table to print the test case, the table name was stored in tableConfig
    numTables = doc.Tables.Count
    for i in range(1, numTables + 1):
        # Check if the table exists in tableConfig
        index = str(i)
        if index in tableConfig:
            # Get table name
            tableName = tableConfig[index]
            # Get table
            table = doc.Tables(i)

            # Open a file to write the test case
            caseFile = os.path.join(root, tableName + '.txt')
            file_object = open(caseFile, 'w+')
            print_case_body(file_object, tableName, table)
            file_object.close()


def main():
    # Parse the command-line operator
    shortargs = 'c:'
    folder = r'C:\Users\31415\Desktop\反馈单.doc'  # directory contains the word document and the configure files
    opts, args = getopt.getopt(sys.argv[1:], shortargs)
    for opt, val in opts:
        if opt == '-d':
            folder = val
            break

    if folder == '':
        print
        "Input a directory which contains the word documents and the config files\n"

    # Traverse all files under given folder and generate test cases for them.
    for root, dirs, files in os.walk(folder):
        for file in files:
            if file.endswith('.docx'):
                wordFile = os.path.join(root, file)
                configFile = wordFile.replace('.docx', '.cfg')
                if os.path.isfile(configFile):
                    generate_Robot_test_cases(root, wordFile, configFile)
                else:
                    print
                    configFile + " not found!\n"


main()