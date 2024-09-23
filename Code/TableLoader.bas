Attribute VB_Name = "TableLoader"
Option Explicit

Public Function LoadTableData(ByVal tableName As String, ws As Worksheet) As IVariantArray2D

     On Error GoTo eh
     
    Dim table As ListObject: Set table = CoreLibrary.WorksheetTable(ws.Name, tableName)
    Dim tableData As Variant: tableData = table.DataBodyRange.value
    Dim dataArray As IVariantArray2D: Set dataArray = CoreLibrary.Construct2DVariantArray_Array(tableData, CoreLibrary.ArrayOptions(1))
    
    Set LoadTableData = dataArray
    Exit Function
eh:
    Set LoadTableData = Nothing

End Function

Public Function FindMatchingRow(array2D As IVariantArray2D, ByVal key As String, ByVal columnIndex As Long) As Long

    Dim row As Long
    Dim rowExtents As CoreLibrary.CArrayExtent: Set rowExtents = array2D.rowExtent
    Dim matchingRow As Long: matchingRow = -1
    For row = rowExtents.LowerBound To rowExtents.UpperBound
        If array2D.Element(CoreLibrary.Array2DIndex(row, columnIndex)) = key Then
            matchingRow = row
            Exit For
        End If
    Next row
    
    FindMatchingRow = matchingRow

End Function

Public Function FindMatchingRows(array2D As IVariantArray2D, ByVal key As String, ByVal columnIndex As Long) As ArrayList

    Dim row As Long
    Dim rowList As New ArrayList
    
    Dim rowExtents As CoreLibrary.CArrayExtent: Set rowExtents = array2D.rowExtent
    For row = rowExtents.LowerBound To rowExtents.UpperBound
        If array2D.Element(CoreLibrary.Array2DIndex(row, columnIndex)) = key Then
            rowList.Add row
        End If
    Next row
    
    Set FindMatchingRows = rowList
    
End Function

Public Function ExtractColumn(data2D As IVariantArray2D, ByVal columnIndex As Long) As ArrayList

    Dim namesList As New ArrayList
    Dim row As Long
    Dim rowExtent As CArrayExtent: Set rowExtent = data2D.rowExtent
    
    For row = rowExtent.LowerBound To rowExtent.UpperBound
        Dim value As String: value = data2D.Element(CoreLibrary.Array2DIndex(row, columnIndex))
        If CoreLibrary.IsValidInput(value) Then
            namesList.Add value
        End If
    Next row
    
    Set ExtractColumn = namesList

End Function

Public Function ExtractRowColumns(data2D As IVariantArray2D, rows As ArrayList, ByVal columnIndex As Long) As ArrayList

    Dim namesList As New ArrayList
    Dim row As Long
    For row = 0 To rows.Count - 1
        Dim value As String: value = data2D.Element(CoreLibrary.Array2DIndex(rows(row), columnIndex))
        If CoreLibrary.IsValidInput(value) Then
            namesList.Add value
        End If
    Next row
    
    Set ExtractRowColumns = namesList

End Function
