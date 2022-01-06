VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "UserForm3"
   ClientHeight    =   4584
   ClientLeft      =   -24
   ClientTop       =   -600
   ClientWidth     =   6744
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

    Dim studentInfo() As String
    Dim columns As Integer
    columns = 5
    
    Dim lastRow As String
    
    ReDim studentInfo(columns)
    
    studentInfo(0) = TextBox1.Value
    studentInfo(1) = TextBox2.Value
    studentInfo(2) = TextBox3.Value
    studentInfo(3) = TextBox4.Value
    studentInfo(4) = TextBox5.Value
    studentInfo(5) = TextBox6.Value
     
    Dim bool As Boolean
    Dim SheetName As String
    SheetName = "Student List"
    
    bool = WorksheetExists(SheetName)
    
    'if the sheet exists, add info from textbox to the list
    'otherwise create the list then add

    If bool = True Then
        For i = 0 To columns
            With ThisWorkbook.Sheets(SheetName)
                lastRow = .UsedRange.Rows(.UsedRange.Rows.Count).Row
                .Range("A" & lastRow).Offset(0, i).Value = studentInfo(i)
            End With
        Next
    Else
        Call UserForm1.studentsquery_Click
        For i = 0 To columns
            With ThisWorkbook.Sheets(SheetName)
                lastRow = .UsedRange.Rows(.UsedRange.Rows.Count).Row
                .Range("A" & lastRow).Offset(0, i).Value = studentInfo(i)
            End With
        Next
    End If
    
    UserForm3.Hide
    
    UserForm1.Show

End Sub

Private Sub Label4_Click()

End Sub

Private Sub Label6_Click()

End Sub

Private Sub TextBox1_Change()
    
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'on close of sub-form, reopen main hub
    If CloseMode = 0 Then
        UserForm3.Hide
        
        UserForm1.Show
    End If
 
End Sub
