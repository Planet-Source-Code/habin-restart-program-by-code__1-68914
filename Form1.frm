VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Restart Program"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "Restart"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************
'Description: This program demonstrates how to restart the programe follow code but not external .
'Author:Habin yellow_river_boy@hotmail.com
'***********************************



Sub RestartMe()
Dim strPath As String
Dim strstrPathName As String

    strPath = App.Path
    
    If Right(strPath, 1) <> "\" Then strPath = strPath + "\"
    
    strPathName = strPath + App.EXEName + ".EXE"
    
    
    Shell strPathName & " " & Chr(34) & " RESTART" & Chr(34), vbHide
    
    End
    
End Sub

Private Sub Command1_Click()
    
    'Restart program
    RestartMe
    
End Sub

Private Sub Form_Load()
    
Dim strCmd As String
    
    
    strCmd = Command()      'Read Command Line
    
    If InStr(strCmd, "RESTART") <= 0 Then
        If App.PrevInstance Then
            MsgBox "Program can not start repeated", vbInformation
            End
        End If
    End If
    
    
    Label1.Caption = "Start Time:" & Now
    
End Sub
