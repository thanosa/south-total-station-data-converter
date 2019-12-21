Attribute VB_Name = "WorkSheetManager"
Const MAX_WORKSHEET_NAMING_LENGTH = 31
    
Public Sub Main(sourceWsName As String)
    
    Dim destWsName As String
    destWsName = GetCopiedWorksheetName(sourceWsName)
    
    Sheets.Add.Name = destWsName
    
    Call TransformData(Worksheets(sourceWsName), Worksheets(destWsName), 1, 1)

End Sub

Private Function GetCopiedWorksheetName(originalWorksheetName As String) As String
    GetCopiedWorksheetName = Left(originalWorksheetName & " " & GetTimestamp, MAX_WORKSHEET_NAMING_LENGTH)
End Function

Private Function GetTimestamp() As String
    GetTimestamp = Format(Now(), "yyyyMMddhhmmss")
End Function


