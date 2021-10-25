import PySimpleGUI as sg

import datetime

import pandas as pd

desired_width = 500
todayDate = datetime.date.today()
twoWeeksOuts = todayDate + datetime.timedelta(days=14)
yearVar = todayDate + datetime.timedelta(days=28835)

print(str(yearVar))

sg.theme('DarkBlack')
file_column = [
    [
        sg.Text("File Grabber"),
        sg.In(size=(25, 1), enable_events=True, key="-FILE GRAB-"),
        sg.FileBrowse(),
    ],
    [
        sg.Multiline(size=(75, 75), key="-LIST ITEMS-"),
    ]
]

new_view_of_file = [
    [sg.Text("Choose a .xlsx file:")],
    [sg.Text(key="-TOUT-")],
    [sg.Text(key="-FILE-")],
]

layout = [
    [
        sg.Column(file_column),
        sg.HSeparator(),
        sg.Column(new_view_of_file),
    ]
]
window = sg.Window("Hello", layout, location=(0, 0), size=(1000, 800)).finalize()

while True:
    event, values = window.read()

    if event == "Exit" or event == sg.WIN_CLOSED:
        break

    if event == "-FILE GRAB-":
        file = values["-FILE GRAB-"]
        file = file
        window["-FILE-"].update(file)
        try:
            workbookGUI = pd.read_excel(file, sheet_name="Sheet1")
            workbook2GUI = pd.read_excel(file, sheet_name="Sheet1")

            customerName = workbookGUI["Customer Name"]

            userName = workbookGUI["Assigned To"]
            sizeOfArray = len(userName)
            print(sizeOfArray)

            pd.to_datetime(workbook2GUI["ETA"], format="%b %d %Y")
            eta = workbook2GUI["ETA"]

            customerNameArr = [""]
            etaArr = [""]
            etaToCompare: datetime

            for x in range(sizeOfArray):
                customerNameArr.append(customerName.iloc[0 + x])
            for x in range(sizeOfArray):
                etaArr.append(eta.iloc[0 + x])

            for x in range(sizeOfArray):
                etaStr = etaArr[1 + x]
                if etaArr[1 + x] == "ERROR":
                    etaArr[1 + x] = yearVar
                else:
                    etaToCompare = etaArr[1 + x]
                if etaToCompare < twoWeeksOuts:
                    customerErrStr = str(x) + "Customer " + customerNameArr[
                        0 + x] + "is showing as an error. Please note down reason for error in the timeline."
                    if etaStr == "ERROR":
                        setText = window["-LIST ITEMS-"]
                        setText.update(setText.get() + "\n" + "\n" + customerErrStr)
                    else:
                        withinTwoWeeks = str(x) + " PROACTIVE CONTACT TASK: Called " + customerNameArr[
                            1 + x] + " to advise of ETA of product for " + str(
                            etaStr) + " and that we are still on schedule. Advised Independent PROvider will contact them within two business days of when product is available for installation."
                        if "00:00:00" in withinTwoWeeks:
                            setText = window["-LIST ITEMS-"]
                            setText.update(setText.get() + "\n" + "\n" + withinTwoWeeks.replace("00:00:00 ", "")
                                           + "\n" + "-------------------------------------------------------------------------------------------------")
                else:
                    if etaStr == "ERROR":
                        setText = window["-LIST ITEMS-"]
                        setText.update(setText.get() + "\n" + "\n" + customerErrStr)
                    else:
                        pastTwoWeeks = str(x) + " PROACTIVE CONTACT TASK: Called " + customerNameArr[
                            0 + x] + " to advise of ETA of product for " + str(
                            etaStr) + " and that we are still on schedule. Promised follow-up within two weeks. Created follow-up task for two weeks on " + twoWeeksOuts.strftime(
                            "%b %d %Y") + ". If you are working follow-up task, please contact customer to advise of current ETA or lead time."
                        if "00:00:00" in pastTwoWeeks:
                            setText = window["-LIST ITEMS-"]
                            setText.update(setText.get() + "\n" + "\n" + pastTwoWeeks.replace("00:00:00 ", "")
                                           + "\n" + "-------------------------------------------------------------------------------------------------")
        except FileNotFoundError:
            window["-FILE-"].update("File not found")
window.close()
