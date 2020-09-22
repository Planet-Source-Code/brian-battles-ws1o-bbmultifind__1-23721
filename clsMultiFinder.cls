VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMultiFinder"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function MultiFind(strStringToSearch As String, strItemsToFind As String) As Integer
   
    '---------------------------------------------------------------
    ' Purpose   : find multiple substrings within a string
    ' Parameters: strStringToSearch, strItemsToFind
    ' Returns   : Number of matches
    ' Modified  : 6/2/2001 By BB WS1O
    '---------------------------------------------------------------

    Dim MyArray() As String
    Dim I         As Integer
    Dim iResult   As Integer
    Dim iAccum    As Integer
    
    On Error GoTo Err_MultiFind
    
    ' start by making sure this is zero
    iAccum = 0
    ' grab the individual string items to search for
    MyArray() = Split(strItemsToFind, ",")
        ' comb through the target string as many
        ' times as we have items to hunt for
        For I = LBound(MyArray) To UBound(MyArray)
            ' make sure this starts as zero each time through
            iResult = 0
            ' see if our first array item is hiding inside the target string
            iResult = InStr(1, strStringToSearch, MyArray(I))
                ' if we get anything higher than zero here, we have a match
                If iResult > 0 Then
                    ' so then let's call this one match
                    iResult = 1
                End If
            ' add this match to our accumulated matches number
            iAccum = iAccum + iResult
        Next I
    MultiFind = iAccum
    
Exit_MultiFind:
    
    On Error Resume Next
    Erase MyArray
    On Error GoTo 0
    Exit Function

Err_MultiFind:
    
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "In clsMultiFinder" & " during " & "MultiFind" & vbCrLf & vbCrLf & Err.Source, vbInformation, App.Title & " ADVISORY"
            Resume Exit_MultiFind
    End Select

End Function
