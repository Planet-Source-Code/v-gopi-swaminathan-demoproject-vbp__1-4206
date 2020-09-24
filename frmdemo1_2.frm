VERSION 5.00
Object = "{02B5E320-7292-11CF-93D5-0020AF99504A}#1.0#0"; "MSCHART.OCX"
Begin VB.Form frmdemo1_2 
   BackColor       =   &H00FFFF80&
   Caption         =   "Form5"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form5"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "View Graph"
      Height          =   495
      Left            =   7560
      TabIndex        =   1
      Top             =   9960
      Width           =   975
   End
   Begin MSChartLib.MSChart MSChart1 
      Height          =   9375
      Left            =   0
      OleObjectBlob   =   "frmdemo1_2.frx":0000
      TabIndex        =   0
      Top             =   120
      Width           =   13695
   End
End
Attribute VB_Name = "frmdemo1_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
 
 
  Dim db1 As Database
  Dim rs1 As Recordset
     
    Set db1 = OpenDatabase("C:\demo\demoproject.mdb")
    Set rs1 = db1.OpenRecordset("demoproject")

    
  
MSChart1.Enabled = True
MSChart1.Visible = True
 
rs1.MoveLast
'Storing all the values of the records in an array arr(m,n)
'where m is the number of records and n is the number of fields
m = rs1.RecordCount
n = rs1.Fields.Count - 1

rs1.MoveFirst
ReDim arr(m, n)
With rs1
        For i = 1 To m
        For j = 1 To n
        If j = 1 Then
        arr(i, j) = !USA
        'val1 = val1 + !Pondy
        'MsgBox (arr(i, j))
        End If
        If j = 2 Then
        arr(i, j) = !japan
        'val2 = val2 + !kkl
        End If
        If j = 3 Then
        arr(i, j) = !germany
        'val3 = val3 + !Mahe
        End If
        If j = 4 Then
        arr(i, j) = !india
        'val4 = val4 + !yanam
        End If
        Next j
        .MoveNext
        Next i
End With
rs1.MoveLast
With rs1
 
End With
rs1.MoveFirst
With MSChart1
        ' Displays a 2d chart with 4 columns and 4 rows
        ' data.
        '.Rowlabel will give label to Rows in X axis
        '.Columnlabel will give label to Columns in Y axis
        .Title.Text = "Country Wise Expenditure on Elections Since 1985"
        .chartType = VtChChartType2dBar
        .ColumnCount = n
        .RowCount = m
        For Row = 1 To m
            For Column = 1 To n
                .Column = Column
                .Row = Row
                If Row = 1 Then
                .RowLabel = "1985"
                End If
                If Row = 2 Then
                .RowLabel = "1990"
                End If
                If Row = 3 Then
                .RowLabel = "1995"
                End If
                If Row = 4 Then
                .RowLabel = "Till Date"
                End If
                If Row = 5 Then
                .RowLabel = "Upto 4:30 p.m."
                End If
                If Row = 6 Then
                .RowLabel = "Upto 6:30 p.m"
                End If
                
                If Column = 1 Then
                .ColumnLabel = "USA"
                End If
                If Column = 2 Then
                .ColumnLabel = "JAPAN"
                End If
                
                If Column = 3 Then
                .ColumnLabel = "GERMANY"
                End If
                If Column = 4 Then
                .ColumnLabel = "INDIA"
                End If
                '.data will accept the values to the datapoints.
                .Data = arr(Row, Column)
            Next Column
        Next Row
        ' Use the chart as the backdrop of the legend.
        .ShowLegend = True
        .SelectPart VtChPartTypePlot, index1, index2, index3, index4
        .EditCopy
        .SelectPart VtChPartTypeLegend, index1, index2, index3, index4
        .EditPaste
    End With
 
    With MSChart1.Plot
    .Axis(VtChAxisIdY).AxisTitle = "Expenditure(In Million Dollars)"

    .Axis(VtChAxisIdX).AxisTitle = "Countries"
    End With
         With MSChart1.Plot
         'Filling colours for each item
         .SeriesCollection.Item(1).DataPoints.Item(-1).Brush.FillColor.Red = 133
         .SeriesCollection.Item(1).DataPoints.Item(-1).Brush.FillColor.Green = 45
         .SeriesCollection.Item(1).DataPoints.Item(-1).Brush.FillColor.Blue = 56
         
         .SeriesCollection.Item(4).DataPoints.Item(-1).Brush.FillColor.Red = 200
         .SeriesCollection.Item(4).DataPoints.Item(-1).Brush.FillColor.Green = 100
         .SeriesCollection.Item(4).DataPoints.Item(-1).Brush.FillColor.Blue = 56
         
         'To display the data values at the desired location over the datapoints of the item
         .SeriesCollection.Item(1).DataPoints.Item(-1).DataPointLabel.LocationType = VtChLabelLocationTypeAbovePoint
         .SeriesCollection.Item(2).DataPoints.Item(-1).DataPointLabel.LocationType = VtChLabelLocationTypeAbovePoint
         .SeriesCollection.Item(3).DataPoints.Item(-1).DataPointLabel.LocationType = VtChLabelLocationTypeAbovePoint
         .SeriesCollection.Item(4).DataPoints.Item(-1).DataPointLabel.LocationType = VtChLabelLocationTypeAbovePoint
        
          
         
    End With

End Sub


 
Private Sub Command2_Click()
End
End Sub
Private Sub Command3_Click()

Form5.Shape1.Visible = False
Form5.Shape2.Visible = False
Form5.Command1.Visible = False
Form5.Command2.Visible = False

 Dim Msg ' Declare variable.
    On Error GoTo ErrorHandler  ' Set up error handler.
    PrintForm   ' Print form.
    Exit Sub
ErrorHandler:
    Msg = "The form can't be printed."
    MsgBox Msg  ' Display message.
    Resume Next
    End Sub

Private Sub Form_Load()
MSChart1.Visible = False
End Sub
