VERSION 5.00
Begin VB.Form frmMultiFindTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "   Test BB's MultiFind DLL"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMultiFindTest 
      Height          =   1470
      Left            =   45
      TabIndex        =   3
      Top             =   -60
      Width           =   6030
      Begin VB.TextBox txtNumberOfMatches 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   345
         Left            =   5325
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   900
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.TextBox txtItemsToFind 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   75
         TabIndex        =   5
         Text            =   "search,string,type,want,you"
         Top             =   1020
         Width           =   4950
      End
      Begin VB.TextBox txtStringToSearch 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   75
         TabIndex        =   4
         Text            =   "This is a test string to search in, type anything you want"
         Top             =   405
         Width           =   4950
      End
      Begin VB.Label lblMatches 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Matches Found"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5235
         TabIndex        =   9
         Top             =   450
         Visible         =   0   'False
         Width           =   660
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblItemsToFind 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter a group of terms to find in the string here"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   825
         TabIndex        =   7
         Top             =   795
         Width           =   3465
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblStringToSearch 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter a string to search in here"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1380
         TabIndex        =   6
         Top             =   150
         Width           =   2340
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear And Do Another"
      Enabled         =   0   'False
      Height          =   360
      Left            =   45
      TabIndex        =   1
      Top             =   1485
      Width           =   1935
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Enough, I'm Done"
      Height          =   360
      Left            =   4125
      TabIndex        =   2
      Top             =   1485
      Width           =   1965
   End
   Begin VB.CommandButton cmdCheckString 
      Caption         =   "Check The String"
      Height          =   360
      Left            =   2070
      TabIndex        =   0
      Top             =   1485
      Width           =   1965
   End
End
Attribute VB_Name = "frmMultiFindTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Module     : frmMultiFindTest
' Description: Simple routine to test BBMultiFind DLL
' Procedures : cmdCheckString_Click()
'              cmdClear_Click()
'              cmdExit_Click()
'              txtItemsToFind_KeyUp(KeyCode As Integer, Shift As Integer)
'              txtStringToSearch_KeyUp(KeyCode As Integer, Shift As Integer)
'
' Modified   : 6/2/2001 by B Battles

Option Explicit
Private Sub cmdCheckString_Click()
    
    ' --------------------------------------------------
    ' Comments  : run the string through the DLL and display the results
    ' Modified  : 6/2/2001 by B Battles
    ' --------------------------------------------------
    
    On Error GoTo Err_cmdCheckString_Click
    
    txtNumberOfMatches = MultiFind(txtStringToSearch, txtItemsToFind)
        If txtNumberOfMatches < 1 Then
            txtNumberOfMatches.ForeColor = &HFF&   ' rayed
        Else
            txtNumberOfMatches.ForeColor = &HFFFF& ' yaller
        End If
    txtNumberOfMatches.Visible = True
    lblMatches.Visible = True
    
Exit_cmdCheckString_Click:
    
    Exit Sub
    
Err_cmdCheckString_Click:
    
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "during cmdCheckString_Click, in frmMultiFindTest", vbInformation, "Advisory"
            Resume Exit_cmdCheckString_Click
    End Select
    
End Sub
Private Sub cmdClear_Click()
    
    ' --------------------------------------------------
    ' Comments  : clear textboxes, hide results
    ' Modified  : 6/2/2001 by B Battles
    ' --------------------------------------------------
    
    On Error GoTo Err_cmdClear_Click
    
    txtItemsToFind = ""
    txtStringToSearch = ""
    txtNumberOfMatches = ""
    txtNumberOfMatches.Visible = False
    lblMatches.Visible = False
    cmdClear.Enabled = False
    
Exit_cmdClear_Click:
    
    Exit Sub
    
Err_cmdClear_Click:
    
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "during cmdClear_Click, in frmMultiFindTest", vbInformation, "Advisory"
            Resume Exit_cmdClear_Click
    End Select
    
End Sub
Private Sub cmdExit_Click()
    
    ' --------------------------------------------------
    ' Comments  : bail
    ' Modified  : 6/2/2001 by B Battles
    ' --------------------------------------------------
    
    On Error GoTo Err_cmdExit_Click
    
    Unload Me
    
Exit_cmdExit_Click:
    
    Exit Sub
    
Err_cmdExit_Click:
    
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "during cmdExit_Click, in frmMultiFindTest", vbInformation, "Advisory"
            Resume Exit_cmdExit_Click
    End Select
    
End Sub
Private Sub txtItemsToFind_KeyUp(KeyCode As Integer, Shift As Integer)
    
    ' --------------------------------------------------
    ' Comments  : enable the Clear button
    ' Parameters: KeyCode, Shift
    ' Modified  : 6/2/2001 by B Battles
    ' --------------------------------------------------
    
    On Error GoTo Err_txtItemsToFind_KeyUp
    
    If txtItemsToFind <> "" Then
        cmdClear.Enabled = True
    End If
    
Exit_txtItemsToFind_KeyUp:
    
    Exit Sub
    
Err_txtItemsToFind_KeyUp:
    
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "during txtItemsToFind_KeyUp, in frmMultiFindTest", vbInformation, "Advisory"
            Resume Exit_txtItemsToFind_KeyUp
    End Select
    
End Sub
Private Sub txtStringToSearch_KeyUp(KeyCode As Integer, Shift As Integer)
    
    ' --------------------------------------------------
    ' Comments  : enable the Clear button
    ' Parameters: KeyCode, Shift
    ' Modified  : 6/2/2001 by B Battles
    ' --------------------------------------------------
    
    On Error GoTo Err_txtStringToSearch_KeyUp
    
    If txtStringToSearch <> "" Then
        cmdClear.Enabled = True
    End If
    
Exit_txtStringToSearch_KeyUp:
    
    Exit Sub
    
Err_txtStringToSearch_KeyUp:
    
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & "during txtStringToSearch_KeyUp, in frmMultiFindTest", vbInformation, "Advisory"
            Resume Exit_txtStringToSearch_KeyUp
    End Select
    
End Sub
