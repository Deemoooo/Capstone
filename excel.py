import os
from ocr import run_ocr
import pandas as pd
import shutil
import win32com.client as win32
import threading
import pythoncom
import cv2
import xlsxwriter

# files = ['Obligor_3_DBS_Bills_Statements_Apr.pdf',
    #          'Obligor_5_Facility_Letter_Dah_Sing.pdf',
    #          'Obligor 1 - CAT spreadsheet']
    # # 'Obligor_7_Citi_GBP_Bank_Statements']

def creator(files, printer):
    printer.sprint('Initializing OCR...', 0, 0)

    if threading.currentThread().getName() != 'MainThread':
        pythoncom.CoInitialize()


    #socket.emit()

    docs = ['Bank Statement', 'Bill Statement', 'Loan Repayment Schedule', 'CAT Template']

    template_loc = "static/template"
    temp_loc = "static/temp"
    output_loc = "static/output"
    upload_loc = "static/uploads"
    wd = os.getcwd().replace("\\", "/") + "/"

    template = template_loc + '/template.xlsm'
    output = output_loc + '/output.xlsm'

    shutil.copy(template, output)

    def adder(file_path):
        xl = win32.gencache.EnsureDispatch('Excel.Application')
        xl.Visible = False  # change to False after development
        xl.DisplayAlerts = False
        xl.AskToUpdateLinks = False
        output_xl = xl.Workbooks.Open(wd + output_loc + '/output.xlsm')
        file_xl = xl.Workbooks.Open(file_path)
        file_xl.Sheets(1).Activate()

        for j in xl.ActiveWorkbook.Sheets:
            xl.DisplayAlerts = False
            j.Copy(After=output_xl.Sheets(output_xl.Sheets.Count))
        xl.DisplayAlerts = False
        file_xl.Close(SaveChanges=False)

        output_xl.Sheets(1).Activate()
        for k in xl.ActiveWorkbook.Sheets:
            if k.Name not in ["Introduction", "Summary", "User Metric"]:
                k.Activate()
                k.Visible = False

        output_xl.Save()

    # def adder(file_path):
    #     xl = win32.gencache.EnsureDispatch('Excel.Application')
    #     xl.Visible = False  # change to False after development
    #     xl.DisplayAlerts = False
    #     output_xl = xl.Workbooks.Open(wd + output_loc + '/output.xlsm')
    #     file_xl = xl.Workbooks.Open(file_path)
    #     file_xl.Sheets(1).Activate()
    #
    #     for j in xl.ActiveWorkbook.Sheets:
    #         j.Activate()
    #         j.Visible = 1
    #         j.Copy(After=output_xl.Sheets(output_xl.Sheets.Count))
    #     file_xl.Close()
    #
    #     output_xl.Sheets(1).Activate()
    #     for k in xl.ActiveWorkbook.Sheets:
    #         if k.Name != "Introduction":
    #             k.Activate()
    #             k.Visible = False
    #
    #     output_xl.Save()
    #     xl.Quit()

    # def adder(file_path):
    #     xl = win32.gencache.EnsureDispatch('Excel.Application')
    #     xl.Visible = False  # change to False after development
    #     xl.DisplayAlerts = False
    #     xl.AskToUpdateLinks = False
    #     output_xl = xl.Workbooks.Open(wd + output_loc + '/output.xlsm', UpdateLinks=False)
    #     file_xl = xl.Workbooks.Open(file_path, UpdateLinks=False)
    #     file_xl.Sheets(1).Activate()
    #
    #     for j in xl.ActiveWorkbook.Sheets:
    #         j.Activate()
    #         j.Visible = 1
    #         j.Copy(After=output_xl.Sheets(output_xl.Sheets.Count))
    #     file_xl.Close()
    #
    #     output_xl.Sheets(1).Activate()
    #     for k in xl.ActiveWorkbook.Sheets:
    #         if k.Name != "Introduction":
    #             k.Activate()
    #             k.Visible = False
    #
    #     output_xl.Save()
    #     xl.Quit()


    def identifier(file):
        if 'bank' in file.lower():
            return 'BaS', docs[0]
        elif 'bill' in file.lower():
            return 'BiS', docs[1]
        elif 'facility' in file.lower():
            return 'LR', docs[2]
        elif 'cat' in file.lower():
            return 'CAT', docs[3]
        else:
            return "Not a compatible document type."

    printer.sprint('Identifying documents...', 4, 0)
    #frames = {'BiS': None, 'LR': None, 'BaS': None, 'CAT': None}
    frames = {}
    for file in files:
        doc_type, doc_name = identifier(file)
        if doc_type == 'CAT':
            printer.sprint('Reading CAT Template...', 7, 0)
            adder(file)
            continue
        data, coor, conf = run_ocr(file, doc_type, printer).multi_page_reader()
        frames[doc_type] = {'filename': file, 'docname': doc_name, 'data': data, 'coor': coor, 'conf': conf}

    printer.sprint('Storing Data...', 11, 0)
    with pd.ExcelWriter(temp_loc + '/temp.xlsx') as writer:
        for doc_type in frames:
            if doc_type == 'CAT':
                continue
            doc_name = frames[doc_type]['docname']
            coor = '_coor' + doc_type
            conf = '_conf' + doc_type
            frames[doc_type]['data'].to_excel(writer, sheet_name=doc_name)
            frames[doc_type]['coor'].to_excel(writer, sheet_name=coor)
            frames[doc_type]['conf'].to_excel(writer, sheet_name=conf)
            frames[doc_type]['sheets'] = [doc_name, coor, conf]

    adder(wd + temp_loc + '/temp.xlsx')

    xl = win32.gencache.EnsureDispatch('Excel.Application')
    xl.Visible = False  # change to False after development
    xl.DisplayAlerts = False
    xl.AskToUpdateLinks = False

    output_xl = xl.Workbooks.Open(wd + output_loc + '/output.xlsm', UpdateLinks=False)
    #output_xl.Sheets("Sheet1").Delete()

    output_xl.Sheets(1).Activate()
    output_xl.VBProject.VBComponents('Sheet1').CodeModule.AddFromString('Option Explicit')

    # try:
    #     xl.Application.Run("output.xlsm!Module3.InsertPivotTable")
    # except:
    #     pass

    _sheets = {}
    for i, j in enumerate(xl.ActiveWorkbook.Sheets):
        j.Activate()
        _sheets[j.Name] = [i + 1, j.CodeName]


    printer.sprint('Creating Excel file...', 10, 0)
    for i in frames:
        if i == "CAT":
            continue
        data_sheet, coor_sheet, conf_sheet = frames[i]['sheets']
        print(data_sheet, coor_sheet, conf_sheet)
        # output_xl.Sheets(data_sheet).Visible = False
        # output_xl.Sheets(coor_sheet).Visible = False
        # output_xl.Sheets(conf_sheet).Visible = False
        # output_xl.Sheets(data_sheet).Activate()
        data_ws = output_xl.Sheets(data_sheet)
        data_ws.Visible = True
        data_ws.Activate()
        data = data_ws.Range("B:Z")
        data.Select()
        data.FormatConditions.Add(2, Formula1='=if({0}!B1 <> "", {0}!B1<85, False)'.format(conf_sheet))
        #data.FormatConditions(data.FormatConditions.Count).SetFirstPriority()
        data.FormatConditions(data.FormatConditions.Count).Interior.ColorIndex = 22
        output_xl.Save()
        pagecol = frames[i]['coor'].columns.get_loc("Page Number")
        col = chr(pagecol + 98).capitalize()
        maxpg = frames[i]['coor']["Page Number"].max()

        pagels = {}
        pageurl = {}
        for j in range(int(maxpg)):
            name = wd + 'sample{}/{}.jpg'.format(i, j + 1)
            img = cv2.imread(name, cv2.IMREAD_UNCHANGED)
            pagels[j + 1] = img.shape
            pageurl[j + 1] = name
        w_factor = 600 / pagels[1][1]

        selector_code = '''Option Explicit
    
    
    Private Sub Worksheet_SelectionChange(ByVal Target As Range)
        On Error Resume Next
        'Dim words As Collection
        'Set words = CollectionCreator()
        Dim Range As Range
        Dim rCell As Range
    
         If Selection.Count > 0 And Selection.Count < 3 Then
            On Error Resume Next
            Call DeleteAllShapes
            For Each rCell In Target.Cells
                Dim rng As String
                Dim val As String
                Dim i As Integer
                Dim msgString As String
                Dim coordinates() As String
    
                rng = rCell.Address(RowAbsolute:=False, ColumnAbsolute:=False)
                val = Worksheets("{0}").Range(rng).Value
    
                    If Not val = "" Then
                        coordinates = Split(val, " ")
                        Dim vArr() As String
                        Dim row_num As String
                        Dim pn_address As String
                        Dim pn As String
                        Dim name As String
                        Dim top As Double
                        Dim shp as Object
    
                        vArr = Split(rCell.Address(True, False), "$")
                        row_num = vArr(1)
                        pn_address = "{1}" & row_num
                        pn = Worksheets("{0}").Range(pn_address).Value
    
                        name = "page" & pn
                        ActiveSheet.Shapes(name).Visible = True
                        ActiveSheet.Shapes(name).Width = 600

                        top = ActiveSheet.Shapes(name).top
    
                        For Each shp In ActiveSheet.Shapes
                           If shp.name <> name Then shp.Visible = False
                        Next shp
    
                    'If Exists(words, range) Then
                        Dim x As Double: x = coordinates(0)*{2}
                        Dim y As Double: y = coordinates(1)*{2} + top
                        Dim w As Double: w = coordinates(2)*{2}
                        Dim h As Double: h = coordinates(3)*{2}
                        'Dim name As String: name = coordinates(4)
                    Call BoxCreator(x, y, w, h)
                    'Else
                        'Call DeleteAllShapes
                    End If
            Next rCell
        End If
    
    End Sub'''.format(coor_sheet, col, w_factor)


        # print(_sheets)
        # for i in output_xl.VBProject.VBComponents:
        #     print(i)
        #     print(i.Name)
        #     print(i.Properties)
        output_xl.VBProject.VBComponents(str(_sheets[data_sheet][1])).CodeModule.AddFromString(selector_code)
        data_ws.Visible = True
        data_ws.Activate()
        xl.Range("B1").Select()  # add the () at the end here
        xl.ActiveWindow.FreezePanes = True
        data_ws.Columns("A").ColumnWidth = 106
        data_ws.Columns("A").Font.ColorIndex = 2
        #xl.Selection.ClearContents()


        # print(pageurl)
        for l in pageurl:
            # data_ws.Activate()
            entry_index = frames[i]['coor'].index[frames[i]['coor']["Page Number"] == l].tolist()[0] + 1
            top = xl.Range("A" + str(entry_index)).Top
            link = pageurl[l].replace("/", "\\")
            pic = data_ws.Shapes.AddPicture(link,
                                            LinkToFile=False, SaveWithDocument=True, Left=0, Top=0, Width=-1, Height=-1)

            pic.Name = 'page' + str(l)
            pic.Left = 0
            pic.Top = top
            pic.Width = 600
            if l == 1:
                pic.Visible = True
            else:
                pic.Visible = False

    output_xl.Sheets(1).Activate()
    xl.DisplayAlerts = False

    output_xl.Save()
    xl.Quit()




