Attribute VB_Name = "TransformDataModule"
Option Explicit

Public Sub TransformData(sourceWs As Worksheet, _
                            destWs As Worksheet, _
                            dataStartAtRow As Integer, _
                            dataStartAtCol As Integer)
    
    Dim KEYS_IGNORE_ARRAY(0 To 4) As String
    KEYS_IGNORE_ARRAY(0) = "JOB"
    KEYS_IGNORE_ARRAY(1) = "INST"
    KEYS_IGNORE_ARRAY(2) = "UNITS"
    KEYS_IGNORE_ARRAY(3) = "XYZ"
    KEYS_IGNORE_ARRAY(4) = "HV"

    Const KEYS_STN = "STN"
    Const KEYS_ST = "ST"
    Const KEYS_BS = "BS"
    Const KEYS_SS = "SS"
    Const KEYS_SD = "SD"

    
    ' Column variables in source
    Dim keyCol As Integer
    keyCol = dataStartAtCol
    Dim dataCol(1 To 3) As Integer
    dataCol(1) = dataStartAtCol + 1
    dataCol(2) = dataStartAtCol + 2
    dataCol(3) = dataStartAtCol + 3
    
    ' Row variable in source
    Dim srcRow As Integer, destRow As Integer
    srcRow = dataStartAtRow
    destRow = 1
    
    ' Key value
    Dim tempKey As String
    
    ' Read variables. The index is the column offset from the key column
    Dim A(1 To 2) As String
    Dim B(1 To 1) As String
    Dim K(1 To 3) As String
    Dim L(1 To 3) As String
    
    
    ' Main algorithm
    Do
        tempKey = sourceWs.Cells(srcRow, keyCol).Value

        If Len(Trim(tempKey)) = 0 Then
            Exit Do
        ElseIf IsInArray(tempKey, KEYS_IGNORE_ARRAY) Then
            srcRow = srcRow + 1
        ElseIf (tempKey = KEYS_STN) Then
            A(1) = sourceWs.Cells(srcRow, dataCol(1)).Value
            A(2) = sourceWs.Cells(srcRow, dataCol(2)).Value
            
            Do
                srcRow = srcRow + 1
                
                tempKey = sourceWs.Cells(srcRow, keyCol).Value
                
                If (tempKey = KEYS_BS) Then
                    B(1) = sourceWs.Cells(srcRow, dataCol(1)).Value
                    
                    Dim tempColA As Integer
                    For tempColA = 1 To 4
                        destWs.Cells(destRow, tempColA).NumberFormat = "@"
                    Next tempColA

                    destWs.Cells(destRow, 1) = "ST"
                    destWs.Cells(destRow, 2) = A(1)
                    destWs.Cells(destRow, 3) = A(2)
                    destWs.Cells(destRow, 4) = B(1)
                    
                    Dim errA As Integer
                    Dim tempCellA As Range
                    For Each tempCellA In Range("A" & destRow & ":D" & destRow).Cells
                        For errA = 1 To 8
                            tempCellA.Errors(errA).Ignore = True
                        Next errA
                    Next tempCellA
                    
                    destRow = destRow + 1
                    
                    Exit Do
                End If
            Loop
        
        ElseIf (tempKey = KEYS_BS) Or (tempKey = KEYS_SS) Then
            K(1) = sourceWs.Cells(srcRow, dataCol(1)).Value
            K(2) = sourceWs.Cells(srcRow, dataCol(2)).Value
            K(3) = sourceWs.Cells(srcRow, dataCol(3)).Value
            
            Do
                srcRow = srcRow + 1
                
                tempKey = sourceWs.Cells(srcRow, keyCol).Value
                
                If (tempKey = KEYS_SD) Then
                    L(1) = sourceWs.Cells(srcRow, dataCol(1)).Value
                    L(2) = sourceWs.Cells(srcRow, dataCol(2)).Value
                    L(3) = sourceWs.Cells(srcRow, dataCol(3)).Value
                    
                    Dim tempColB As Integer
                    For tempColB = 1 To 7
                        destWs.Cells(destRow, tempColB).NumberFormat = "@"
                    Next tempColB
                    
                    destWs.Cells(destRow, 1) = "SS"
                    destWs.Cells(destRow, 2) = K(1)
                    destWs.Cells(destRow, 3) = L(1)
                    destWs.Cells(destRow, 4) = L(2)
                    destWs.Cells(destRow, 5) = L(3)
                    destWs.Cells(destRow, 6) = K(2)
                    destWs.Cells(destRow, 7) = K(3)
                    
                    Dim errB As Integer
                    Dim tempCellB As Range
                    For Each tempCellB In Range("A" & destRow & ":G" & destRow).Cells
                        For errB = 1 To 8
                            tempCellB.Errors(errB).Ignore = True
                        Next errB
                    Next tempCellB
                    
                    destRow = destRow + 1
                    
                    srcRow = srcRow + 1
                    
                    Exit Do
                End If
            Loop
        Else
            srcRow = srcRow + 1
        End If
    
    Loop
    
End Sub

Private Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    ' Checks if the array contains a specific value
    IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function


