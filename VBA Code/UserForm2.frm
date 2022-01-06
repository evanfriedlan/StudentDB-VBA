VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   4008
   ClientLeft      =   36
   ClientTop       =   168
   ClientWidth     =   3984
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()

End Sub

Private Sub ComboBox2_Change()
    
    
    
End Sub

Private Sub CommandButton1_Click()

     
    Dim location As String
    Dim SQL As String
    
    Dim rs As New ADODB.Recordset
    Dim courseCode As String
    Dim numStudents As String
    Dim course As Variant
    Dim term As Variant
    
    course = ComboBox1.Value
    term = ComboBox2.Value
    
    location = "\Students.accdb"
    
    SQL = "SELECT Courses.[Course Title], Courses.[Course Code], CRN.CRN " & _
    "FROM (Courses INNER JOIN CRN ON Courses.[Course ID] = CRN.CourseID) " & _
    "WHERE CRN.CourseID = " & course & " AND CRN.TermDesc = '" & term & "'"

    Set rs = DBLoad(location, SQL)

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
    Dim data2() As String
    Dim fields2() As String
    Dim columns2 As Integer
    columns2 = 3
    
    ReDim fields2(columns2) As String
    

    fields2(0) = "ID"
    fields2(1) = "StudentID"
    fields2(2) = "CRN"
    fields2(3) = "Final Grade"
    


    
    SQL2 = "SELECT * FROM Grades WHERE CRN = " & data(0, 2)
    


    Set rs2 = DBLoad(location, SQL2)

    rc = 0

    
    data2 = DBToArray(rs2, fields2, columns2)
    
        
    With rs2
        Do Until .EOF
            rc = rc + 1
            .MoveNext
        Loop
    End With
    
    rs2.Close
    
    Sheets.Add

    For i = 0 To UBound(fields)
        ThisWorkbook.ActiveSheet.Range("A1").Offset(0, i).Value = fields(i)
    Next
    
    ThisWorkbook.ActiveSheet.Range("D1").Value = "No. Of Students"
    ThisWorkbook.ActiveSheet.Range("E1").Value = "Course ID"
    ThisWorkbook.ActiveSheet.Range("F1").Value = "Term"

    For i = 0 To UBound(data)
        For k = 0 To 2
            ThisWorkbook.ActiveSheet.Range("A1").Offset(1 + i, k).Value = data(i, k)
        Next
        ThisWorkbook.ActiveSheet.Range("A1").Offset(1 + i, 3).Value = rc
        ThisWorkbook.ActiveSheet.Range("A1").Offset(1 + i, 4).Value = course
        ThisWorkbook.ActiveSheet.Range("A1").Offset(1 + i, 5).Value = term
    Next
    
        ThisWorkbook.ActiveSheet.Range("A1").Offset(3, 0).Value = fields2(0)
        ThisWorkbook.ActiveSheet.Range("A1").Offset(3, 1).Value = fields2(1)
        ThisWorkbook.ActiveSheet.Range("A1").Offset(3, 2).Value = fields2(2)
        ThisWorkbook.ActiveSheet.Range("A1").Offset(3, 3).Value = fields2(3)
    If rc > 0 Then
    
        
    
        For i = 0 To UBound(data2)
           For k = 0 To 3
               ThisWorkbook.ActiveSheet.Range("A1").Offset(4 + i, k).Value = data2(i, k)
           Next
        Next
    
        Dim mean As Long
        Dim median As Long
        Dim mode As Long
        Dim StdDev As Long
        Dim lastRow As Long
       
        lastRow = Range("C" & Rows.Count).End(xlUp).Row
    
        Dim numRange As Range
        Set numRange = ThisWorkbook.ActiveSheet.Range("D5:D" & lastRow)
        mean = Application.WorksheetFunction.Average(numRange)
        median = Application.WorksheetFunction.median(numRange)
        mode = Application.WorksheetFunction.Mode_Sngl(numRange)
        StdDev = Application.WorksheetFunction.StDev_P(numRange)
         
        
        With ThisWorkbook.ActiveSheet
            .Range("I1").Value = "Class Mean"
            .Range("I2").Value = "Class Median"
            .Range("I3").Value = "Class Mode"
            .Range("I4").Value = "Class Standard Deviation"
            .Range("J1").Value = mean
            .Range("J2").Value = median
            .Range("J3").Value = mode
            .Range("J4").Value = StdDev
        End With
 
 
        ThisWorkbook.ActiveSheet.UsedRange.columns.AutoFit

    
    
        Dim cht As ChartObject
        Dim rng As Range
    
        Set rng = ActiveSheet.Range("D5:D" & lastRow)
    
        Set cht = ActiveSheet.ChartObjects.Add(Left:=400, Width:=450, _
        Top:=100, Height:=250)
       
        cht.Chart.SetSourceData Source:=rng, PlotBy:=xlColumns
        
        

     Else
        For k = 0 To 3
            ThisWorkbook.ActiveSheet.Range("A1").Offset(4, k).Value = 0
        Next
        
        ThisWorkbook.ActiveSheet.UsedRange.columns.AutoFit

     End If
  
    

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label2_Click()

End Sub
  
Private Sub UserForm_Initialize()
    'adding the options by course Id and term to the box
    ComboBox1.AddItem 2
    ComboBox1.AddItem 3
    ComboBox1.AddItem 4
    ComboBox1.AddItem 5
    ComboBox1.AddItem 6
    ComboBox1.AddItem 7
    
    ComboBox2.AddItem "Fall"
    ComboBox2.AddItem "Winter"
    ComboBox2.AddItem "Spring"
End Sub
