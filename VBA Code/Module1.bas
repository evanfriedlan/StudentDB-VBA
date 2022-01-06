Attribute VB_Name = "Module1"
Option Explicit
' ==== CP212 Windows Application Programming ===============+
' Name: Evan Friedlan
' Student ID: 180922230
' Date: April 9, 2021
' Program title: Assignment 5
' Description: Student Database Option
'===========================================================+

Sub UserForm3_Terminate()
    UserForm1.Show
    'close Sub-form for button 6a, going back to main menu
End Sub

Sub UserForm()

    UserForm1.Show

End Sub

Function DBLoad(location As String, SQL As String) As ADODB.Recordset
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset

    'open database provided it is located in the workbook's folder

    With cn
        .ConnectionString = "Data Source=" & ThisWorkbook.Path & location
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With

    rs.Open SQL, cn, adOpenDynamic, adLockOptimistic

    Set DBLoad = rs
    'send recordsheet to be parsed in various ways
    

End Function

Function WorksheetExists(SheetName As String) As Boolean
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SheetName)
    On Error GoTo 0
    
    'simple boolean to see if a worksheet has been created yet
    
    If ws Is Nothing Then
        WorksheetExists = False
    Else
        WorksheetExists = True
    End If
End Function

Function DBToArray(rs As ADODB.Recordset, fields() As String, columns As Integer)
    Dim data() As String
    Dim i As Integer
    Dim rc As Long

    rc = 0
    
    
    With rs
        Do Until .EOF
            rc = rc + 1
            .MoveNext
        Loop  ' counting how many rows are in the recordset
        If rc > 0 Then
            .MoveFirst 'return to top so we can find the actual values next
        End If
    End With
    
    If rc > 0 Then
       ReDim data(rc - 1, columns) 'shape 2d array to fit all incoming data
       
       rc = 0

       With rs
           Do Until .EOF
               For i = 0 To columns
                    data(rc, i) = .fields(fields(i))
               Next
               rc = rc + 1
               .MoveNext
           Loop
           .MoveFirst
       End With
    Else
        'blank array is returned if no records are available for error-avoidance
        ReDim data(0, 0)
        data(0, 0) = " "
    End If
    
    
    DBToArray = data
    
End Function

Public Sub NewSheet(SheetName As String)

    Sheets.Add.Name = SheetName

    With Sheets(SheetName)
        .Range("A1").Value = "Student List"
        .Range("A1").Style = "Title"
        .Rows(2).Font.Bold = True
        .Rows(2).HorizontalAlignment = xlCenter
        .Range("A1:H10").columns.AutoFit
    End With
    
End Sub



Function WriteToDBEnrolments(data() As String, rs As ADODB.Recordset)
'unique function for button 3, writes from .dat file to Enrolment in the students.accdb
    
    Dim tester As Variant
    Dim i As Integer
    Dim curID As Variant
    Dim curSID As Variant
    Dim curCRN As Variant

    'There will be 1 of 2 errors on run of this (button 3).
    'The records are still added to the database, however

    On Error GoTo pass

    With rs
       For i = 0 To UBound(data())
           curID = i + 159
           curSID = data(i, 0)
           curCRN = data(i, 1)
           
           .AddNew
           .fields("ID") = curID
           .fields("StudentID") = curSID
           .fields("CRN") = curCRN
       Next
           .Update
    End With
    
pass:
    MsgBox "The enrolment records were added, please check your database"
       
End Function

Function TxtFileToArray(delimiter As String)
    Dim fd As Office.FileDialog
    Dim file As String
    Dim filename As Variant
    Dim dataLine As String
    Dim notCancel As Boolean
    Dim dataIn() As String
    Dim dataOut() As String
    
    Dim i As Integer, j As Integer, p As Integer, q As Integer, r As Integer
    Dim k As Variant

    Dim letter As String
    Dim nextletter As String
    Dim cats As Integer
    Dim temp() As String
    Dim columns As Integer
    
    cats = 0
    i = 0
    Set fd = Application.FileDialog(msoFileDialogOpen)
    
    'open file picker and select ONE(1) file
    
    With fd
        notCancel = .Show
        If notCancel Then
            filename = .SelectedItems(1)
        End If
    End With

    'on close, read text file using delimiter to split each line into values for array
        
    If notCancel Then
        file = filename


        Open file For Input As #1
    
        Do Until EOF(1)
            ReDim Preserve dataIn(i + 1)
            Line Input #1, dataLine
            dataIn(i) = dataLine
            i = i + 1
        Loop

        
        Close #1
        
        For r = 1 To Len(dataIn(0))
            letter = Mid(dataIn(0), r, 1)
            nextletter = Mid(dataIn(0), r + 1, 1)
           
            If letter = delimiter Then
                If Not nextletter = delimiter Then
                    cats = cats + 1
                End If
            End If
        Next
        
        
        
        ReDim dataOut(i, cats) As String
        
        p = 0
        Do Until p = i
            ReDim temp(cats)
            temp = Split(dataIn(p), delimiter)

            q = 0
            
            Do Until q = (cats + 1)
                dataOut(p, q) = temp(q)
                q = q + 1
            Loop
            
            p = p + 1
        Loop
        TxtFileToArray = dataOut
    Else
        'more fail case protection
        Dim fail(0) As String
        TxtFileToArray = fail
    End If
    

End Function
