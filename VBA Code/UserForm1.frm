VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   6564
   ClientLeft      =   -1344
   ClientTop       =   -9780.001
   ClientWidth     =   4800
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub addstudent_Click()
    
    UserForm1.Hide
    
    UserForm3.Show
    
    
End Sub

Private Sub classenrollments_Click()

    Dim location As String
    Dim SQL As String
    
    Dim rs As New ADODB.Recordset
    Dim course As Variant
    Dim term As Variant
    Dim courseCode As String
    Dim coursetitle As String
    Dim numStudents As String
    
    Dim booler As Boolean
    booler = False
    
    
    'enforce user to input one of the database's possible selections
    
    Do Until booler = True
        course = InputBox("Enter Course By ID (2 - 7)")
        
        Select Case course
        Case 2 To 7
            booler = True
        Case Else
            booler = False
        End Select
    Loop
    
    booler = False
    
    Do Until booler = True
        term = InputBox("Enter Term (Fall, Spring, Winter) [Case Sensitive]")
        
        Select Case term
        Case "Fall", "Spring", "Winter"
            booler = True
        Case Else
            booler = False
        End Select
    Loop
        
   
    location = "\Students.accdb"
    
    SQL = "SELECT Courses.[Course Title], Courses.[Course Code], CRN.CRN " & _
    "FROM (Courses INNER JOIN CRN ON Courses.[Course ID] = CRN.CourseID) " & _
    "WHERE CRN.CourseID = " & course & " AND CRN.TermDesc = '" & term & "'"

    Set rs = DBLoad(location, SQL)
    'Calling function to create the recordset
    

    Dim data() As String
    Dim fields() As String
    Dim columns As Integer
    Dim i As Integer, k As Integer
    columns = 2
    
    ReDim fields(columns) As String
    
    fields(0) = "Course Code"
    fields(1) = "Course Title"
    fields(2) = "CRN"

    
    data = DBToArray(rs, fields, columns)
    'this function parses the recordset

    rs.Close
    
    'time for second recordset, using the information gleaned from the first
    Dim rs2 As New ADODB.Recordset
    Dim rc As Long
    Dim SQL2 As String
    
    SQL2 = "SELECT * FROM Enrolments WHERE CRN = " & data(0, 2) 'needed the CRN to count

    Set rs2 = DBLoad(location, SQL2)

    rc = 0
    
    'counting # of enrolments in the given class
    
    With rs2
        Do Until .EOF
            rc = rc + 1
            .MoveNext
        Loop
    End With
    

    Sheets.Add

    For i = 0 To UBound(fields)
        ThisWorkbook.ActiveSheet.Range("A1").Offset(0, i).Value = fields(i)
    Next
    
    ThisWorkbook.ActiveSheet.Range("D1").Value = "No. Of Students"
    ThisWorkbook.ActiveSheet.Range("E1").Value = "Course ID"
    ThisWorkbook.ActiveSheet.Range("F1").Value = "Term"
    
    'output to new worksheet
    

    For i = 0 To UBound(data)
        For k = 0 To 2
            ThisWorkbook.ActiveSheet.Range("A1").Offset(1 + i, k).Value = data(i, k)
        Next
        ThisWorkbook.ActiveSheet.Range("A1").Offset(1 + i, 3).Value = rc
        ThisWorkbook.ActiveSheet.Range("A1").Offset(1 + i, 4).Value = course
        ThisWorkbook.ActiveSheet.Range("A1").Offset(1 + i, 5).Value = term
    Next
    
    ThisWorkbook.ActiveSheet.UsedRange.columns.AutoFit
    

End Sub

Private Sub findbycity_Click()

'loading and parsing database to array for a given city

    Dim dbTable As String
    Dim location As String
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim cityName As String
    
    cityName = InputBox("Enter city name")
    dbTable = "Students"
    location = "\Students.accdb"
    SQL = "SELECT * FROM " & dbTable & " WHERE City = '" & cityName & "'"
    
    Set rs = DBLoad(location, SQL)

    Dim data() As String
    Dim fields() As String
    Dim columns As Integer
    Dim i As Integer, k As Integer
    columns = 1
    
    ReDim fields(columns) As String
    
    fields(0) = "First Name"
    fields(1) = "Last Name"
    
    data = DBToArray(rs, fields, columns)
    
    Sheets.Add
    
'writing said array to new worksheet
    
    For i = 0 To UBound(fields)
        ThisWorkbook.ActiveSheet.Range("A1").Offset(0, i).Value = fields(i)
    Next
    ThisWorkbook.ActiveSheet.Range("C1").Value = "City"
    
    Debug.Print UBound(data)

    If UBound(data) > 1 Then
        For i = 0 To UBound(data)
            For k = 0 To 1
                ThisWorkbook.ActiveSheet.Range("A1").Offset(1 + i, k).Value = data(i, k)
            Next
            ThisWorkbook.ActiveSheet.Range("A1").Offset(1 + i, 2).Value = cityName
        Next
    End If
    
    ThisWorkbook.ActiveSheet.UsedRange.columns.AutoFit
    


End Sub

Private Sub importstudents_Click()

    Dim data() As String
    Dim i As Integer, q As Integer, k As Integer
    Dim delimiter As String
    Dim fields(9) As String
    
    fields(0) = "StudentID"
    fields(1) = "LastName"
    fields(2) = "FirstName"
    fields(3) = "EmailAddress"
    fields(4) = "Level"
    fields(5) = "DateOfBirth"
    fields(6) = "HomePhone"
    fields(7) = "Address"
    fields(8) = "City"
    fields(9) = "StateProvince"
    
    delimiter = ","

'define variables for parsing dat file to array

    data = TxtFileToArray(delimiter)
  
' then write array to worksheet with titles


    If UBound(data) > 0 Then
        Sheets.Add
        For k = 0 To 9
            ThisWorkbook.ActiveSheet.Range("A1").Offset(0, k).Value = fields(k)
        Next
        
        For i = 0 To UBound(data())
            For q = 0 To 10
                ThisWorkbook.ActiveSheet.Range("A1").Offset(1 + i, q).Value = data(i, q)
            Next
        Next
    End If

    ThisWorkbook.ActiveSheet.UsedRange.columns.AutoFit
    
  
End Sub

Private Sub stjacobs_Click()

    Dim dbTable As String
    Dim location As String
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim cityName As String
    
'simple Query to count the number of students from a given city
    
    cityName = "St. Jacobs"
    dbTable = "Students"
    location = "\Students.accdb"
    SQL = "SELECT * FROM " & dbTable & " WHERE City = '" & cityName & "'"
    
    Set rs = DBLoad(location, SQL)

    Dim data() As String
    Dim fields() As String
    Dim columns As Integer
    Dim i As Integer, k As Integer
    columns = 1
    
    ReDim fields(columns) As String
    
    fields(0) = "First Name"
    fields(1) = "Last Name"
    
    data = DBToArray(rs, fields, columns)
    
    MsgBox "There are " & UBound(data) & " students from " & cityName
    

End Sub
 
Private Sub studentspercity_Click()
    
' very similar to above sub, with customizable city name

    Dim dbTable As String
    Dim location As String
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim cityName As String
    
    
    cityName = InputBox("Enter city name")
    dbTable = "Students"
    location = "\Students.accdb"
    SQL = "SELECT * FROM " & dbTable & " WHERE City = '" & cityName & "'"
    
    Set rs = DBLoad(location, SQL)

    Dim data() As String
    Dim fields() As String
    Dim columns As Integer
    Dim i As Integer, k As Integer
    columns = 1
    
    ReDim fields(columns) As String
    
    fields(0) = "First Name"
    fields(1) = "Last Name"
    
    data = DBToArray(rs, fields, columns)
    
    MsgBox "There are " & UBound(data) & " students from " & cityName



End Sub

Sub studentsquery_Click()
    Dim bool As Boolean
    Dim SheetName As String
    SheetName = "Student List"
    
    bool = WorksheetExists(SheetName) 'checking if the worksheet has been created
    
    If bool = True Then
        ThisWorkbook.Sheets(SheetName).Delete
    End If
    
    Dim dbTable As String
    Dim location As String
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    
    dbTable = "Students"
    location = "\Students.accdb"
    SQL = "SELECT * FROM " & dbTable
    
    Set rs = DBLoad(location, SQL)
 
    Dim data() As String
    Dim fields() As String
    Dim columns As Integer
    Dim i As Integer, k As Integer
    columns = 3
    
    ReDim fields(columns) As String
    
    fields(0) = "First Name"
    fields(1) = "Last Name"
    fields(2) = "E-mail Address"
    fields(3) = "City"
    
    data = DBToArray(rs, fields, columns)

    Call NewSheet(SheetName)
    
    ' create new formatted sheet

    'print to formatted worksheet
    
    For i = 0 To UBound(fields)
        ThisWorkbook.Sheets("Student List").Range("A1").Offset(1, i).Value = fields(i)
    Next

    For i = 0 To UBound(data)
        For k = 0 To 3
            ThisWorkbook.Sheets("Student List").Range("A1").Offset(2 + i, k).Value = data(i, k)
        Next
    Next
    
    ThisWorkbook.Sheets("Student List").UsedRange.columns.AutoFit

End Sub

Sub wordcreate_Click()

    UserForm1.Hide
    UserForm2.Show
  


End Sub

Private Sub finalgrades_Click()
'reads data from dat file and stores it into an array, not printed or used
    
    Dim data() As String
    Dim i As Integer, q As Integer
    Dim delimiter As String
    
    delimiter = vbTab
    
    data = TxtFileToArray(delimiter)
    
    If UBound(data) > 0 Then MsgBox "finalgrades.dat has been stored to an array"
    
    
End Sub

Sub writetoenrolments_Click()

    Dim data() As String
    Dim i As Integer, k As Integer
    Dim delimiter As String
    Dim rs As New ADODB.Recordset

    delimiter = vbTab

    data = TxtFileToArray(delimiter)
    
    
    Debug.Print UBound(data)
    
    If UBound(data) > 0 Then
        Dim dbTable As String
        Dim location As String
        Dim SQL As String
        
        dbTable = "Enrolments"
        location = "\Students.accdb"
        SQL = "SELECT * FROM " & dbTable
        
        Set rs = DBLoad(location, SQL)
        
        Call WriteToDBEnrolments(data, rs)
    End If
        
End Sub
