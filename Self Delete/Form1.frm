VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   Caption         =   "Self Delete Example"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Self Delete This Executable"
      Height          =   975
      Left            =   1335
      TabIndex        =   0
      Top             =   1178
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If InDesignMode = False Then
    Dim j As String
    Dim g As String
    
    j = NameFix(App.EXEName) & ".exe" ' Name Fix for name such Copy Of Project1.exe
    
    Filenum = FreeFile
    
    g = App.Path & "testx.bat"
    
    j = "del " & j & vbCrLf & _
    "del testx.bat"
    
    Open g For Output As Filenum
    Print #Filenum, j
    Close Filenum
    
    Shell g, vbHide                  'Call our batch file and Quit
    End
Else
    MsgBox "Connot delete Exe in VB Design Mode Please make a exe first", vbExclamation, "Error Deleting EXE"
End If
End Sub

Public Function NameFix(Txt As String) As String
Dim temp As String
Dim j
j = Split(Txt, " ", -1, 1)
If UBound(j, 1) <> 0 Then
    For i = 0 To UBound(j, 1)
    temp = temp & j(i)
    Next
    
        If Len(temp) < 7 Then
        temp = temp & "~1"
        Else
        temp = Left(temp, 6) & "~1"
        End If
        NameFix = temp
Else
    NameFix = Txt
End If
End Function

Public Function InDesignMode() As Boolean
    On Error GoTo Err
    Debug.Print 1 / 0
    InDesignMode = False
    Exit Function
Err:
    InDesignMode = True
End Function

