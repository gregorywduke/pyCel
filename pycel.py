# pyCel GUI

import PySimpleGUI as sg
import openpyxl

from openpyxl.styles.fills import PatternFill
from openpyxl.styles import Font, colors
from newXL import *

# GUI layout
layout = [[sg.Text('Please select your task')],
          [sg.Button('Recap')], [sg.Button('Scheduling')],
          [sg.Button('Exit')]]

# Create window
window = sg.Window('pyCel', layout)

# Event Loop
while True:
    event, values = window.read()

    if event == sg.WIN_CLOSED:
        break
    if event == 'Exit':
        break

    if event == 'Recap':
        # Get data file
        filename = sg.popup_get_file('Select Data File')
        template = sg.popup_get_file('Select Template File')
        logo = sg.popup_get_file('Select Logo File')
        #output = sg.popup_get_folder('Select Folder to Output New Files')
        weekNum = sg.popup_get_text('Insert Week Number (FORMAT: "WK # " - Please put space after #)')
        dateRange = sg.popup_get_text('Insert Date Range (FORMAT: "MONTH/DAY - MONTH/DAY")')

        row = 2
        clNum = 1
        fileNum = 0

        # Get number of rows as upper bound
        rowMax = getRows(filename)
        rowMax += 1

        # Creates one file per call, runs thru entire dataset
        while row < rowMax:
        #while row < 10:
            row = recap(filename, template, row, clNum, weekNum, dateRange, logo )
            fileNum += 1
            clNum += 1

        filesText = " files created."
        filesProcessed = f'{fileNum}{filesText}'
        sg.popup(filesProcessed)

    if event == 'Scheduling':
        sg.popup('Not implemented yet!')

        """# Get data file
        filename = sg.popup_get_file('Select Excel File to Process')
        # output = sg.popup_get_folder('Select Folder to Output New Files')
        weekNum = sg.popup_get_text('Insert Week Number')

        row = 2
        clNum = 1
        fileNum = 0

        # Get number of rows as upper bound
        rowMax = getRows(filename)
        rowMax += 1

        # while row < rowMax:
        # while row < 20:
        # row = buildFile(filename, row, clNum)
        fileNum += 1
        # clNum += 1

        filesText = " files created."
        filesProcessed = f'{fileNum}{filesText}'
        sg.popup(filesProcessed)"""

window.close()




