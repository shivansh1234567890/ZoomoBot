import re
import os


def isValidHexaCode(str):
    # Regex to check valid
    # hexadecimal color code.
    regex = "^#([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$"

    # Compile the ReGex
    p = re.compile(regex)

    # If the string is empty
    # return false
    if(str == None):
        return False

    # Return if the string
    # matched the ReGex
    if(re.search(p, str)):
        return True
    else:
        return False


if __name__ == '__main__':
    # Check if file with name Cache.txt exists
    if os.path.isfile('Cache.txt'):
        # If file exists, delete it
        os.remove('Cache.txt')

    # Create file with name Cache.txt
    name = input('Enter your name (Enter -1 to skip): ')
    while not name:
        name = input('Enter your name (Enter -1 to skip): ')
    if name == '-1':
        name = ''
    skip_column = input('Enter column number to skip: ')
    while (not skip_column or not skip_column.isdigit()) and skip_column != '-1':
        if not skip_column.isdigit():
            print("Invalid input :(\nEnter an integer")
            skip_column = ""
        skip_column = input(
            'Enter column number to skip: ')
    if skip_column == '-1':
        skip_column = ''
    skip_row = input('Enter row number to skip: ')
    while (not skip_row or not skip_row.isdigit()) and skip_row != '-1':
        if not skip_row.isdigit():
            print("Invalid input :(\nEnter an integer")
            skip_row = ""
        skip_row = input('Enter row number to skip: ')
    if skip_row == '-1':
        skip_row = ''
    # Write name, skip_column and skip_row to Cache.txt
    color = input(
        'Enter hex value of the color with which you want to color the cell (Enter -1 to skip): ')
    while not isValidHexaCode(color):
        color = input(
            'Enter hex value of the color with which you want to color the cell (Enter -1 to skip): ')
    if color == '-1':
        color = ''
    with open('Cache.txt', 'w') as f:
        f.write(name + "|Name" + '\n' + skip_column +
                "|No. of columns to skip" + '\n' + skip_row + '|No. of rows to skip\n'+color+'|Cell Color')
