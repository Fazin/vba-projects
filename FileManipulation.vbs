Sub ef_abrir()
    '
    ' Open file in xls save as xlsx and delete old file
    '

    '
    Dim ExcelFile As Workbook


    ano = DatePart("YYYY", Date)

    mes_texto = MonthName(DatePart("M", Date), False)
    mes_num = DatePart("M", Date)
    origem_dados = ActiveWorkbook.Name
    Windows(origem_dados).Activate
    myDateText = Format(Date - 1, "ddmm")

    pathname = "c:\Trabalho\Transporte\" & ano & "\0.-EF Diaria\" & mes_num & ".-" & mes_texto

    Filename = "RF AGRUP " & myDateText & " 3.xls"

    fname = "RF AGRUP " & myDateText & " 3.xlsx"

    Set ExcelFile = Application.Workbooks.Open(pathname & "\" & Filename, local:=True, notify:=False, UpdateLinks:=False)
        With ExcelFile
            .SaveAs pathname & "\" & fname, 51
        End With

        DeleteFile (pathname & "\" & Filename)
        
    Windows(fname).Activate

    arrumar_agrup_ef

    agrup_file_ef
        
    Windows(origem_dados).Activate

End Sub

Function FileExists(ByVal FileToTest As String) As Boolean
'Check if file exists
   FileExists = (Dir(FileToTest) <> "")
End Function

Sub DeleteFile(ByVal FileToDelete As String)
   If FileExists(FileToDelete) Then 'See above
      ' First remove readonly attribute, if set
      SetAttr FileToDelete, vbNormal
      ' Then delete the file
      Kill FileToDelete
   End If
End Sub

Sub ClearFilters()

  'To Clear All Fitlers use the ShowAllData method for
  'for the sheet.  Add error handling to bypass error if
  'no filters are applied.  Does not work for Tables.
  On Error Resume Next
    Sheet1.ShowAllData
  On Error GoTo 0
  

End Sub

Sub Reset()
    Dim pt As PivotTable
    Dim ws As Worksheet
        
    Application.ScreenUpdating = False
    'RefreshSlicersOnWorksheet ActiveSheet
    For Each ws In ActiveWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
            If pt.SourceData = "Base_EF" Then
                pt.SourceData = "Base_EF"
                pt.Update
            End If
        
        Next pt
    Next ws
        

    Application.ScreenUpdating = True
End Sub

Public Sub RefreshSlicersOnWorksheet(ws As Worksheet)
    Dim sc As SlicerCache
    Dim scs As SlicerCaches
    Dim slice As Slicer

    Set scs = ws.Parent.SlicerCaches

    If Not scs Is Nothing Then
        For Each sc In scs
            For Each slice In sc.Slicers
                If slice.Shape.Parent Is ws Then
                    sc.ClearManualFilter
                    Exit For 'unnecessary to check the other slicers of the slicer cache
                End If
            Next slice
        Next sc
    End If

End Sub

Public Sub lastrow_Excel()
    Set sht = ActiveSheet
    sht.Select
    lastrow = sht.Cells(sht.Rows.Count, "A").End(xlUp).Row
End Sub

Public Sub Select_FirstCellFiltered()
    Range([t2], Cells(Rows.Count, "t")).SpecialCells(xlCellTypeVisible)(1).Select 'First of column T with header at the first line
End Sub

Public Sub PivotManipulation()
'Determine the data range you want to pivot
  SrcData = sht.Name & "!" & Range("A1:x" & lastrow).Address(ReferenceStyle:=xlR1C1)

'Create a new worksheet
  Set sht = Sheets.Add

'Where do you want Pivot Table to start?
  StartPvt = sht.Name & "!" & sht.Range("A3").Address(ReferenceStyle:=xlR1C1)

'Create Pivot Cache from Source Data
  Set pvtCache = ActiveWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=SrcData)

'Create Pivot table from Pivot Cache
  Set pvt = pvtCache.CreatePivotTable( _
    TableDestination:=StartPvt, _
    TableName:="Pivot Table Name")

End Sub


Sub BreakExternalLinks()
    'PURPOSE: Breaks all external links that would show up in Excel's "Edit Links" Dialog Box
    'SOURCE: www.TheSpreadsheetGuru.com/The-Code-Vault

    Dim ExternalLinks As Variant
    Dim wb As Workbook
    Dim x As Long

    Set wb = ActiveWorkbook

    'Create an Array of all External Links stored in Workbook
    ExternalLinks = wb.LinkSources(Type:=xlLinkTypeExcelLinks)

    'Loop Through each External Link in ActiveWorkbook and Break it
    For x = 1 To UBound(ExternalLinks)
        wb.BreakLink Name:=ExternalLinks(x), Type:=xlLinkTypeExcelLinks
    Next x

End Sub
