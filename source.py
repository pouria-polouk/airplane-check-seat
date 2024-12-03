
# Designed and programmed by: Pouria Polouk (pouria.polouk@gmail.com)

import flet
import datetime
import os
import sys
import openpyxl
import pandas
import numpy
from openpyxl.styles import Font
import psutil

def main(screen:flet.Page):
    screen.window.width = 700
    screen.window.height = 600
    screen.window.resizable = False
    screen.window.maximizable = False
    screen.title = "Swiss Flight Seat Check"

    excelFileName = ""
    dicCompanyRowAndSeat = {}
    allUniqueCompanyName = set({})
    dropdownCompanyName:flet.Dropdown
    checkExcelFileIsOpened = False

    # ================================================================================== Show all seats and read/write seat problem
    def showAllRowsSeats(dropdownValueCompanyName):

        def mouseOutOver(e):
            if e.data == "true":
                e.control.scale = 1.3
            else:
                e.control.scale = 1
            screen.update()


        def checkRowAndSeat(r, s):
            if len(dicCompanyRowAndSeat) != 0:
                for val in dicCompanyRowAndSeat[dropdownValueCompanyName]:
                    if str(val).strip() != 'nan':
                        rowNumber = ''.join([ch for ch in val if ch.isdigit()])
                        seatAlphabet = ''.join([ch for ch in val if ch.isalpha()])
                        if r == rowNumber.strip() and s == seatAlphabet.strip():
                            return True
            return False

        def tickSeatProblem(e):
            if e.control.style.bgcolor == "transparent":
                e.control.style = flet.ButtonStyle(bgcolor="red", shape=flet.CircleBorder())
                dicCompanyRowAndSeat[dropdownValueCompanyName].append(e.control.data)
            else:
                e.control.style = flet.ButtonStyle(bgcolor="transparent", shape=flet.CircleBorder())
                dicCompanyRowAndSeat[dropdownValueCompanyName].remove(e.control.data)
            screen.update()


        columnRowsSeats.controls.clear()
        btnSaveChangeToFile = flet.ElevatedButton("Save Change(s)", bgcolor='#44df00', color='white', on_click=lambda e: seatFlightClass.changeAddDeleteToExcelFile(screen, excelFileName, dicCompanyRowAndSeat, dropdownCompanyName, "CHANGE_SEAT", "Seat Check"))
        columnRowsSeats.controls.append(btnSaveChangeToFile)

        for r in range(1,46):
            labelRowSeat = flet.Text(value='Row'+str(r), width=50)
            rowHolder = flet.Row(spacing=0)
            rowHolder.controls.append(labelRowSeat)
            for s in range(65,75):
                allButtonStyleRed = flet.ButtonStyle(bgcolor="red", shape=flet.CircleBorder())
                allButtonStyleTransparent = flet.ButtonStyle(bgcolor="transparent", shape=flet.CircleBorder())
                if checkRowAndSeat(str(r), chr(s)):
                    btnRowSeat = flet.ElevatedButton(text=chr(s), color='white', style=allButtonStyleRed, scale=flet.transform.Scale(1), animate_scale=flet.animation.Animation(70, flet.AnimationCurve.BOUNCE_OUT), data=str(r)+chr(s), on_hover=mouseOutOver, on_click=tickSeatProblem)
                else:
                    btnRowSeat = flet.ElevatedButton(text=chr(s), color='white', style=allButtonStyleTransparent, scale=flet.transform.Scale(1), animate_scale=flet.animation.Animation(70, flet.AnimationCurve.BOUNCE_OUT), data=str(r)+chr(s), on_hover=mouseOutOver, on_click=tickSeatProblem)
                rowHolder.controls.append(btnRowSeat)

            columnRowsSeats.controls.append(rowHolder)

    # ================================================================================== Set/Select dropdown menu
    def setDropdownCompany():
        nonlocal dropdownCompanyName
        def selectDropdownCompany(e):
            if dropdownCompanyName.value is not None or dropdownCompanyName.value != "":
                showAllRowsSeats(dropdownCompanyName.value)
                screen.update()

        columnDropdown.controls.clear()
        columnRowsSeats.controls.clear()

        dropdownCompanyName = flet.Dropdown(width=300, options=[flet.dropdown.Option(cmpy) for cmpy in sorted(allUniqueCompanyName)], on_change=selectDropdownCompany)
        dropdownCompanyName.hint_text = "Select Company Name..."

        columnDropdown.controls.append(dropdownCompanyName)
        screen.update()

    #================================================================================== Choose excel file
    def excelFilePickerDialog(exFile:flet.FilePickerResultEvent):
        nonlocal excelFileName
        nonlocal allUniqueCompanyName
        nonlocal dicCompanyRowAndSeat
        nonlocal checkExcelFileIsOpened

        if exFile.files:
            excelFileName = str(exFile.files[0].path)
            screen.title = "Swiss Flight Seat Check" + " ==> " + excelFileName
            screen.update()
            dicCompanyRowAndSeat.clear()
            allUniqueCompanyName.clear()
            checkExcelFileIsOpened = True


            dfSwissAirplane = pandas.read_excel(excelFileName, header=None, skiprows=1)

            if not dfSwissAirplane.empty:
                dfSwissAirplane = dfSwissAirplane.add_prefix("info_", axis=1)
                allUniqueCompanyName = set(dfSwissAirplane['info_0'].unique())

                for companyName in allUniqueCompanyName:
                    valueEachCompanyName = (dfSwissAirplane[dfSwissAirplane['info_0'] == companyName]).values
                    for row in valueEachCompanyName:
                        key = row[0]
                        valueDic = [item for item in row[1:] if item is not numpy.nan]
                        if key in dicCompanyRowAndSeat:
                            dicCompanyRowAndSeat[key].extend(valueDic)
                        else:
                            dicCompanyRowAndSeat[key] = valueDic

            setDropdownCompany()

    # ================================================================================== NavigationBar
    def actionDialog(status):

        action:flet.AlertDialog
        if status == "ADD_COMPANY":
            textCompanyInput = flet.TextField(hint_text="Enter company name")
            action = flet.AlertDialog(title=flet.Text("Add Company"),
                                      content=textCompanyInput,
                                      actions=[flet.ElevatedButton(text="Add", on_click=lambda e: actionFunc(status, textCompanyInput.value)),
                                               flet.ElevatedButton(text="Cancel", on_click=lambda e: actionCloseFunc())]
                                      )
        else:
            action = flet.AlertDialog(title=flet.Text("Delete Company"),
                                      content=dropdownCompanyName,
                                      actions=[flet.ElevatedButton(text="Delete", on_click=lambda e: actionFunc(status, dropdownCompanyName.value)),
                                               flet.ElevatedButton(text="Cancel",on_click=lambda e: actionCloseFunc())]
                                      )

        screen.overlay.append(action)
        action.open = True
        screen.update()

        def actionFunc(status, textValueOrDropdownValue):
            if status == "ADD_COMPANY":
                upperKeyList = [upperKey.upper() for upperKey in dicCompanyRowAndSeat.keys()]
                if textValueOrDropdownValue.strip().upper() in upperKeyList:
                    seatFlightClass.showSnackbar(screen, "Error : Company Exists.")
                elif textValueOrDropdownValue.strip() != "":
                    if seatFlightClass.checkIfExcelfileIsOpened(excelFileName) == False:
                        actionCloseFunc()
                        seatFlightClass.changeAddDeleteToExcelFile(screen, excelFileName, dicCompanyRowAndSeat, dropdownCompanyName, status, textValueOrDropdownValue)
                    else:
                        seatFlightClass.showSnackbar(screen, "Error: The Excel file must exist, be closed, and have write permissions.")
            else:
                if textValueOrDropdownValue in dicCompanyRowAndSeat:
                    if seatFlightClass.checkIfExcelfileIsOpened(excelFileName) == False:
                        columnRowsSeats.controls.clear()
                        actionCloseFunc()
                        seatFlightClass.changeAddDeleteToExcelFile(screen, excelFileName, dicCompanyRowAndSeat, dropdownCompanyName, status, textValueOrDropdownValue)
                    else:
                        seatFlightClass.showSnackbar(screen, "Error: The Excel file must exist, be closed, and have write permissions.")
            screen.update()

        def actionCloseFunc():
            action.open = False
            screen.update()


    def selectedNavItem(e):
        if e.control.selected_index == 0: # Select Excel File...
            excelFilePicker.pick_files(allow_multiple=False, allowed_extensions=['xlsx'])
        elif e.control.selected_index == 1: # Add Company
            if checkExcelFileIsOpened == True:
                actionDialog("ADD_COMPANY")
            else:
                seatFlightClass.showSnackbar(screen, "Warning: At first you must select an Excel file...")
        elif e.control.selected_index == 2: # Delete Company
            if checkExcelFileIsOpened == True:
                actionDialog("DELETE_COMPANY")
            else:
                seatFlightClass.showSnackbar(screen, "Warning: At first you must select an Excel file...")
        elif e.control.selected_index == 3: # About
            contactAction = flet.AlertDialog(title=flet.Text("About"),
                                      content=flet.Text("Designed and programmed by: Pouria Polouk\nGmail: pouria.polouk@gmail.com\n***************************************\nProject Consultant: Soheil Pourkohan"),
                                      actions=[flet.ElevatedButton(text="Ok", on_click=lambda e: showAbout())]
                                      )
            screen.overlay.append(contactAction)
            contactAction.open = True
            screen.update()

        def showAbout():
            contactAction.open = False
            screen.update()


    nav = flet.NavigationBar(
        destinations=[flet.NavigationBarDestination(label='Excel File...', icon=flet.Icons.IMPORT_EXPORT),
                      flet.NavigationBarDestination(label='Add Company', icon=flet.Icons.ADD),
                      flet.NavigationBarDestination(label='Delete Company', icon=flet.Icons.DELETE),
                      flet.NavigationBarDestination(label='Contact', icon=flet.Icons.CONTACTS)],
        bgcolor='white',
        on_change=selectedNavItem
    )
    screen.navigation_bar = nav

    #===========================================================================================

    def resource_path(relative_path):
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(base_path, relative_path)



    excelFilePicker = flet.FilePicker(on_result=excelFilePickerDialog)
    columnDropdown = flet.Column()
    columnRowsSeats = flet.Column(scroll=flet.ScrollMode.AUTO)

    container = flet.Container(content=flet.Column([excelFilePicker, columnDropdown, columnRowsSeats], scroll=flet.ScrollMode.AUTO),
                               padding=20,
                               width=screen.width,
                               height=screen.height - 200,
                               #image_src=os.getcwd() + "\\airplane.jpg",
                               image_src = resource_path("airplane.jpg"),
                               image_fit=flet.ImageFit.COVER)

    screen.add(container)

#================================================================================== Class
class seatFlightClass:

    @classmethod
    def changeAddDeleteToExcelFile(cls, screen, excelFileName, dicCompanyRowAndSeat, dropdownCompanyName, typeOfAction, textValueOrDropdownValue):

        rollBackDicCompanyRowAndSeat = dicCompanyRowAndSeat.copy()
        rollBackDropdownItems = [item for item in dropdownCompanyName.options]

        if typeOfAction == "ADD_COMPANY":
            dicCompanyRowAndSeat[textValueOrDropdownValue.strip()] = [''] # At least included one item
            dropdownCompanyName.options.append(flet.dropdown.Option(textValueOrDropdownValue.strip()))
        elif typeOfAction == "DELETE_COMPANY":
            del dicCompanyRowAndSeat[textValueOrDropdownValue]
            dropdownCompanyName.options = [item for item in dropdownCompanyName.options if item.key != textValueOrDropdownValue]
            dropdownCompanyName.update()

        processIncompelete = False
        for k, v in dicCompanyRowAndSeat.items():
            valuesWithoutEmpty = [item for item in v if str(item).strip() != 'nan' and str(item).strip() != '']
            print(valuesWithoutEmpty)
            if len(valuesWithoutEmpty) > 0:
                if cls.checkCellValueBeforeSave(valuesWithoutEmpty) == True:
                    dicCompanyRowAndSeat[k] = sorted(valuesWithoutEmpty, key=lambda item:(int(item[:-1]), item[-1]))
                else:
                    processIncompelete = True

        allCompany = []
        allSeat = []
        for key, value in dicCompanyRowAndSeat.items():
            i = 0
            while i < len(value):
                allCompany.append(key)
                allSeat.append(value[i: i + 9])
                i += 9

        newChangeDataframe = pandas.DataFrame(allSeat)
        newChangeDataframe.insert(0, 'CompanyName', allCompany)
        newChangeDataframe.set_index('CompanyName', inplace=True)

        textForSaveChange = "Successful: The changes were successful."
        try:
            newChangeDataframe.to_excel(excelFileName, startrow=1, header=False)
            workbook = openpyxl.load_workbook(excelFileName)
            airplaneSheet = workbook.active

            cell = airplaneSheet.cell(row=1, column=2, value=str(datetime.datetime.now().date()))
            cell.font = Font(bold=True, size=15)
            workbook.save(excelFileName)
        except:
            dicCompanyRowAndSeat = rollBackDicCompanyRowAndSeat
            dropdownCompanyName.options = rollBackDropdownItems
            dropdownCompanyName.update()
            textForSaveChange = "Error: The Excel file may be unavailable or open."

        if processIncompelete == True:
            cls.showSnackbar(screen, "Incomplete! invalid value has been entered in the Excel file.")
        else:
            cls.showSnackbar(screen, textForSaveChange)

    @classmethod
    def showSnackbar(cls, screen, msg):
        snack = flet.SnackBar(flet.Text(msg))
        screen.overlay.append(snack)
        snack.open = True
        screen.update()


    @classmethod
    def checkIfExcelfileIsOpened(cls, excelFileName):
        for process in psutil.process_iter(['pid', 'name']):
            if process.info['name'] == 'EXCEL.EXE':
                for file in process.open_files():
                    if file.path == excelFileName:
                        return True
        return False

    @classmethod
    def checkCellValueBeforeSave(cls, valuesWithoutEmpty):
        for value in valuesWithoutEmpty:
            print(value)
            letter = ''.join([l for l in value if l.isdigit()])
            number = ''.join([n for n in value if n.isalpha()])
            #print(letter, number)
            if str(letter).strip() == '' or str(number).strip() == '':
                return False
        return True


#================================================================================= Start Project
flet.app(target=main)
