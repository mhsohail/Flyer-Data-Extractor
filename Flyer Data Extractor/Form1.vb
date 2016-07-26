Imports System.IO
Imports System.Text.RegularExpressions
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports Newtonsoft.Json

Public Class FrmFlyerDataExtrarctor
    Private Sub BtnExtractData_Click(sender As Object, e As EventArgs) Handles BtnExtractData.Click
        Try
            LblInfo.Text = "Extracting data..."
            Dim websiteUrl = TxtWebsiteUrl.Text
            Dim sourceString As String = New System.Net.WebClient().DownloadString(websiteUrl)
            Dim regex = New Regex("window\[\'flyerData\'\] = (\{.+\});")
            Dim match = regex.Match(sourceString)
            If match.Success Then
                Dim value = match.Groups(1).Value
                Dim flyerData = JsonConvert.DeserializeObject(Of FlyerData)(value)
                LblInfo.Text = flyerData.Items.Count.ToString() + " products found." + If(flyerData.Items.Count > 0, " Saving to excel in progress...", "")
                PutResultsInExcel(flyerData.Items)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub PutResultsInExcel(items As List(Of Item))
        'create excel file to save results in
        'save results to excel file
        Dim FilePath = "Products.xlsx"
        If Not (System.IO.File.Exists(FilePath)) Then
            CreateSpreadsheetWorkbook(FilePath)
        End If

        Using fs As FileStream = New FileStream(FilePath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite)
            Using spreadSheetDoc As SpreadsheetDocument = SpreadsheetDocument.Open(fs, True)
                'put headings in first row
                InsertText(spreadSheetDoc, "Name", Convert.ToChar(65 + 0).ToString(), 1)
                InsertText(spreadSheetDoc, "Price", Convert.ToChar(65 + 1).ToString(), 1)
                InsertText(spreadSheetDoc, "URL", Convert.ToChar(65 + 2).ToString(), 1)

                Dim Row As UInt32 = 2
                'foreach (var CalculatedMsa in CalculatedMsas)
                '// if there Is only one address for an MSA, this value will be null
                If items IsNot Nothing Then
                    For Each Item As Item In items
                        InsertText(spreadSheetDoc, Item.display_name, Convert.ToChar(65 + 0).ToString(), Row)
                        InsertText(spreadSheetDoc, Item.current_price, Convert.ToChar(65 + 1).ToString(), Row)
                        InsertText(spreadSheetDoc, Item.url, Convert.ToChar(65 + 2).ToString(), Row)
                        Row = Row + 1
                    Next
                    LblInfo.Text = "Saving to excel complete."
                End If
                'Next

                spreadSheetDoc.Close()
            End Using
        End Using

    End Sub

    Private Sub CreateSpreadsheetWorkbook(filepath As String)
        ' Create a spreadsheet document by supplying the filepath.
        'By default, AutoSave = true, Editable = true, And Type = xlsx.
        Dim spreadsheetDoc = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook)

        'Add a WorkbookPart to the document.
        Dim WorkbookPartObj = spreadsheetDoc.AddWorkbookPart()
        WorkbookPartObj.Workbook = New Workbook()

        'Add a WorksheetPart to the WorkbookPart.
        Dim WorksheetPartObj = WorkbookPartObj.AddNewPart(Of WorksheetPart)()
        WorksheetPartObj.Worksheet = New Worksheet(New SheetData())

        'Add Sheets to the Workbook.
        Dim SheetsObj = spreadsheetDoc.WorkbookPart.Workbook.
        AppendChild(Of Sheets)(New Sheets())

        'Append a New worksheet And associate it with the workbook.
        Dim SheetObj = New Sheet
        With SheetObj
            .Id = spreadsheetDoc.WorkbookPart.GetIdOfPart(WorksheetPartObj)
            .SheetId = 1
            .Name = "mySheet"
        End With
        SheetsObj.Append(SheetObj)
        WorkbookPartObj.Workbook.Save()

        'Close the document
        spreadsheetDoc.Close()
    End Sub

    Private Sub InsertText(spreadSheetObj As SpreadsheetDocument, text As String, Row As String, Column As UInt32)
        'Get the SharedStringTablePart. If it does Not exist, create a New one.
        Dim shareStringPart As SharedStringTablePart
        If spreadSheetObj.WorkbookPart.GetPartsOfType(Of SharedStringTablePart)().Count() > 0 Then
            shareStringPart = spreadSheetObj.WorkbookPart.GetPartsOfType(Of SharedStringTablePart)().First()
        Else
            shareStringPart = spreadSheetObj.WorkbookPart.AddNewPart(Of SharedStringTablePart)()
        End If

        'Insert the text into the SharedStringTablePart.
        Dim Index As Int32 = InsertSharedStringItem(text, shareStringPart)

        'Insert a New worksheet.
        Dim WorksheetPartObj As WorksheetPart = InsertWorksheet(spreadSheetObj.WorkbookPart)

        'Insert cell A1 into the New worksheet.
        Dim CellObj As Cell = InsertCellInWorksheet(Row, Column, WorksheetPartObj)

        'Set the value of cell A1.
        CellObj.CellValue = New CellValue(Index.ToString())
        CellObj.DataType = New EnumValue(Of CellValues)(CellValues.SharedString)

        'Save the New worksheet.
        WorksheetPartObj.Worksheet.Save()
    End Sub

    'Given text And a SharedStringTablePart, creates a SharedStringItem with the specified text 
    'And inserts it into the SharedStringTablePart. If the item already exists, returns its index.
    Private Function InsertSharedStringItem(text As String, shareStringPart As SharedStringTablePart) As Int32
        'If the part does Not contain a SharedStringTable, create one.
        If shareStringPart.SharedStringTable Is Nothing Then
            shareStringPart.SharedStringTable = New SharedStringTable()
        End If
        Dim i As Int32 = 0

        'Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
        For Each ItemObj As SharedStringItem In shareStringPart.SharedStringTable.Elements(Of SharedStringItem)()
            If ItemObj.InnerText = text Then
                Return i
            End If
            i = i + 1
        Next

        'The text does Not exist in the part. Create the SharedStringItem And return its index.
        shareStringPart.SharedStringTable.AppendChild(New SharedStringItem(New DocumentFormat.OpenXml.Spreadsheet.Text(text)))
        shareStringPart.SharedStringTable.Save()

        Return i
    End Function

    'Given a WorkbookPart, inserts a New worksheet.
    Private Function InsertWorksheet(workbookPartObj As WorkbookPart) As WorksheetPart
        'We need single sheet only, if there Is a sheet, return
        If workbookPartObj.WorksheetParts.Count() > 0 Then
            Return workbookPartObj.WorksheetParts.FirstOrDefault()
        End If

        'Add a New worksheet part to the workbook.
        Dim newWorksheetPart As WorksheetPart = workbookPartObj.AddNewPart(Of WorksheetPart)()
        newWorksheetPart.Worksheet = New Worksheet(New SheetData())
        newWorksheetPart.Worksheet.Save()

        Dim SheetsObj As Sheets = workbookPartObj.Workbook.GetFirstChild(Of Sheets)()
        Dim relationshipId As String = workbookPartObj.GetIdOfPart(newWorksheetPart)

        'Get a unique ID for the New sheet.
        Dim SheetId As UInt32 = 1
        If SheetsObj.Elements(Of Sheet)().Count() > 0 Then
            SheetId = SheetsObj.Elements(Of Sheet)().Select(Function(s) s.SheetId.Value).Max() + 1
        End If

        Dim SheetName As String = "Sheet" + SheetId

        'Append the New worksheet And associate it with the workbook.
        Dim SheetObj As Sheet = New Sheet()
        With SheetObj
            .Id = relationshipId
            .SheetId = SheetId
            .Name = SheetName
        End With

        SheetsObj.Append(SheetObj)
        workbookPartObj.Workbook.Save()

        Return newWorksheetPart
    End Function

    'Given a column name, a row index, And a WorksheetPart, inserts a cell into the worksheet. 
    'If the cell already exists, returns it. 
    Private Function InsertCellInWorksheet(columnName As String, rowIndex As UInt32, worksheetPartObj As WorksheetPart) As Cell
        Dim WorksheetObj As Worksheet = worksheetPartObj.Worksheet
        Dim SheetDataObj As SheetData = WorksheetObj.GetFirstChild(Of SheetData)
        Dim cellReference As String = columnName + rowIndex.ToString()

        'If the worksheet does Not contain a row with the specified row index, insert one.
        Dim Row As DocumentFormat.OpenXml.Spreadsheet.Row
        If Not SheetDataObj.Elements(Of DocumentFormat.OpenXml.Spreadsheet.Row)().Where(Function(r) r.RowIndex.Value = rowIndex).Count() = 0 Then
            Row = SheetDataObj.Elements(Of DocumentFormat.OpenXml.Spreadsheet.Row)().Where(Function(r) r.RowIndex.Value = rowIndex).First()
        Else
            Row = New DocumentFormat.OpenXml.Spreadsheet.Row()
            With Row
                .RowIndex = rowIndex
            End With
            SheetDataObj.Append(Row)
        End If

        'If there Is Not a cell with the specified column name, insert one.  
        If Row.Elements(Of Cell)().Where(Function(c) c.CellReference.Value = columnName + rowIndex.ToString()).Count() > 0 Then
            Return Row.Elements(Of Cell)().Where(Function(c) c.CellReference.Value = cellReference).First()
        Else
            'Cells must be in sequential order according to CellReference. Determine where to insert the New cell.
            Dim refCell As Cell = Nothing
            For Each CellObj As Cell In Row.Elements(Of Cell)()
                If String.Compare(CellObj.CellReference.Value, cellReference, True) > 0 Then
                    refCell = CellObj
                    Exit For
                End If
            Next

            Dim NewCell As Cell = New Cell()
            With NewCell
                .CellReference = cellReference
            End With

            Row.InsertBefore(NewCell, refCell)

            WorksheetObj.Save()
            Return NewCell
        End If
    End Function

    Private Sub FrmFlyerDataExtrarctor_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LblInfo.Text = String.Empty
    End Sub
End Class
