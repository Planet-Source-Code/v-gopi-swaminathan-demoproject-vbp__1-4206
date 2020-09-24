VERSION 5.00
Object = "{02B5E320-7292-11CF-93D5-0020AF99504A}#1.0#0"; "MSCHART.OCX"
Begin VB.Form Frmdemo1_1 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "View Graph"
      Height          =   855
      Left            =   6720
      TabIndex        =   1
      Top             =   9720
      Width           =   2535
   End
   Begin MSChartLib.MSChart MSChart1 
      Height          =   8535
      Left            =   960
      OleObjectBlob   =   "frmdemo1_1.frx":0000
      TabIndex        =   0
      Top             =   480
      Width           =   12615
   End
End
Attribute VB_Name = "Frmdemo1_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
    Dim db1 As Database
    Dim rs1 As Recordset
     
   

Private Sub Command1_Click()
 
     MSChart1.Visible = True
     Call disp_chart
   
End Sub

 

 

Private Sub Form_Load()
'To insert
  MSChart1.Visible = False
  MSChart1.Enabled = True
  'The following sample database contains demoproject table for display of graph
  Set db1 = OpenDatabase("c:\demo\demoproject.mdb")
  Set rs1 = db1.OpenRecordset("demoproject")
End Sub

'To close


Private Sub Form_Unload(Cancel As Integer)
 rs1.Close
 db1.Close
 Set rs1 = Nothing
 Set db1 = Nothing
' frmgmain.Show
End Sub
Private Sub disp_chart()
 
 
With rs1
'To display the first record data by bar chart
.MoveFirst
 value1 = Format(!USA, "###.##")
 value2 = Format(!japan, "###.##")
 value3 = Format(!germany, "###.##")
 value4 = Format(!india, "###.##")
 
End With

  
 MSChart1.Visible = True
 With MSChart1
         'To choose the type of the Chart
         .chartType = VtChChartType2dBar
         'A Title given to the Chart
         .TitleText = " Countrywise Expenditure on Elections in billion dollars"
         'Specifying the title location
         .Title.Location.LocationType = VtChLocationTypeTop
         'Determining the number of columns and rows
         
            .RowCount = 1
            .ColumnCount = 4
         'One record is one data series.
         ' Each data series is a collection of columns
         'Each data is assigned to a particular row and column with .data
         
            For Row = 1 To 1
            For Column = 1 To 4
           
           .Column = Column
           .Row = Row
           If Column = 1 Then
        
                .Data = value1
                .ColumnLabel = "USA"
            End If
            If Column = 2 Then
                .Data = value2
                .ColumnLabel = "JAPAN"
            End If
            If Column = 3 Then
              .Data = value3
              .ColumnLabel = "GERMANY"
            End If
            If Column = 4 Then
                .Data = value4
                .ColumnLabel = "INDIA"
            End If
            Next Column
            Next Row
        
        ' Use the chart as the backdrop of the legend.
        .ShowLegend = True
 
        End With
        With MSChart1.Plot
        'To customize the datapoints fillcolors
         .SeriesCollection.Item(1).DataPoints.Item(-1).Brush.FillColor.Red = 200
         .SeriesCollection.Item(1).DataPoints.Item(-1).Brush.FillColor.Green = 50
         .SeriesCollection.Item(1).DataPoints.Item(-1).Brush.FillColor.Blue = 200
         .SeriesCollection.Item(4).DataPoints.Item(-1).Brush.FillColor.Red = 200
         .SeriesCollection.Item(4).DataPoints.Item(-1).Brush.FillColor.Green = 100
         .SeriesCollection.Item(4).DataPoints.Item(-1).Brush.FillColor.Blue = 56
        'To Print the Label Value at the top of the bar of each data
        .SeriesCollection.Item(1).DataPoints.Item(-1).DataPointLabel.LocationType = VtChLabelLocationTypeOutside
        .SeriesCollection.Item(2).DataPoints.Item(-1).DataPointLabel.LocationType = VtChLabelLocationTypeOutside
        .SeriesCollection.Item(3).DataPoints.Item(-1).DataPointLabel.LocationType = VtChLabelLocationTypeOutside
        .SeriesCollection.Item(4).DataPoints.Item(-1).DataPointLabel.LocationType = VtChLabelLocationTypeOutside
        .Axis(VtChAxisIdX).AxisTitle = "COUNTRIES"
        .Axis(VtChAxisIdY).AxisTitle = "EXPENDITURE(In Billion Dollars)"
       
    End With
End Sub

