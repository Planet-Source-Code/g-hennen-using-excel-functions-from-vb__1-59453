VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form1 
   Caption         =   "Using Excel Functions..."
   ClientHeight    =   7305
   ClientLeft      =   5325
   ClientTop       =   2205
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   6480
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   495
      Left            =   4440
      TabIndex        =   5
      Top             =   6720
      Width           =   1935
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   4095
      Left            =   120
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   4
      Top             =   2640
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.CommandButton cmdExponentialCurve 
      Caption         =   "&Exponential Curve Fit"
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      ToolTipText     =   "Calulate and Display Exponential Curve Fit using Excel LogEst Function"
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton cmdStatistics 
      Caption         =   "Data &Statistics"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Calulate and Display Data Statistics using Excel Functions"
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton cmdGetLine 
      Caption         =   "&Line Fit"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      ToolTipText     =   "Calulate and Display Line Fit using Excel LinEst Function"
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   1575
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":1C87
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Using Excel functions in Visual Basic
'/*************************************/
'/* Author: G. Hennen March 2005
'/* garyh@wfeca.net
'/*************************************/
'Excel exposes many powererful functions that can easily be accessed by VB.
'This program demonstrates how to access some Excel functions by passing data arrays
'Thanks to...
'Author: Frank Kusluski (4/26/02)
'for help getting started

Option Explicit

Public mobjExcel As Excel.Application

'arrays to hold actual X & Y data points
Dim mdX(1 To 6) As Double
Dim mdY(1 To 6) As Double
'array to hold calculated regression data points
Dim mdY2(1 To 6) As Double

Private Sub cmdClose_Click()
    Unload Form1
    Set Form1 = Nothing
End Sub

Private Sub cmdExponentialCurve_Click()
    'This routine demonstrates how to use the Excel function LogEst for exponential curve fit
    'LogEst requires data to be passed in an array and it returns the results in a multidimensional array.
    'this routine demonstrates how to pass arrays to an Excel function and retrive the results from the returned multidimensional array
    'using the Excel function Index.
    
    'Regression output values
Dim dSlope As Double
Dim dY_Intercept As Double
Dim dR2 As Double
Dim ix As Integer
    
    'get line equation values
    'pass x & y data array values to Excel LogEst function to get Slope, Y Intercept and Rsquare value.
    'The LogEst function returns an array of regression statistics values representing the line (curve). See LogEst help for array regression statistics values
    
    With mobjExcel.WorksheetFunction
        'Slope (m)
        dSlope = .Index(.LogEst(mdY, mdX), 1, 1)
        'y-intercept (b)
        dY_Intercept = .Index(.LogEst(mdY, mdX), 1, 2)
        'R square. The coefficient of determination. Compares estimated and actual y-values, and ranges in value from 0 to 1.
        'the closer R square is to 1 the better the calculated line "fits" the data
        dR2 = .Index(.LogEst(mdY, mdX, , True), 3, 1)
    End With
    
    Text1.Text = "Y = " & Format(dY_Intercept, "#,###.##0") & " * " & Format(dSlope, "#,###.##0") & "^x" & vbCrLf
    Text1.Text = Text1.Text & "Rsquare = " & Format(dR2, "0.##0")
    
    'load calculated data points for exponential curve into array using the formula Y= b*m^x  (Y= dY_intercept * dSlope^x)
    For ix = 1 To 6
        mdY2(ix) = dY_Intercept * dSlope ^ ix
    Next
    
    'display graph of both lines
    Call SetGraph

End Sub

Private Sub cmdGetLine_Click()
    'This routine demonstrates how to use the Excel function for linear estimate (LinEst)
    'LinEst requires data to be passed in an array and it returns the results in a multidimensional array.
    'So...this routine demonstrates how to pass arrays to an Excel function and retrive the results from the return multidimensional array.
    
    'Regression output values
Dim dSlope As Double
Dim dY_Intercept As Double
Dim dR2 As Double
Dim ix As Integer
    
    'get line equation values
    'pass x & y data array values to Excel LinEst function to get Slope, Y Intercept and Rsquare value.
    'The LinEst function returns an array of regression statistics values representing the line. See LineEst help for array regression statistics values
    'Use Excel's Index function to get the values from the array
    
    With mobjExcel.WorksheetFunction
        'Slope (m)
        dSlope = .Index(.LinEst(mdY, mdX, , True), 1, 1)
        'y-intercept (b)
        dY_Intercept = .Index(.LinEst(mdY, mdX, , True), 1, 2)
        'R square. The coefficient of determination. Compares estimated and actual y-values, and ranges in value from 0 to 1.
        'the closer R square is to 1 the better the calculated line "fits" the data
        dR2 = .Index(.LinEst(mdY, mdX, , True), 3, 1)
    End With
    
    Text1.Text = "Y = " & Format(dSlope, "#,###.##0") & "x + " & Format(dY_Intercept, "#,###.##0") & vbCrLf
    Text1.Text = Text1.Text & "Rsquare = " & Format(dR2, "0.##0")
    
    'load calculated data points for line into array using the formula Y= m*x+b  (Y= dSlope * X + dY_intercept)
    For ix = 1 To 6
        mdY2(ix) = dSlope * ix + dY_Intercept
    Next
    
    'display graph of both lines
    Call SetGraph
End Sub

Private Sub cmdStatistics_Click()
    'data statistics Minimum, Maximum, Median, Average
    'using mdY() array data values
    
    With mobjExcel.WorksheetFunction
        'count
        Text1.Text = "Count = " & .Count(mdY) & vbCrLf
        'Minimum
        Text1.Text = Text1.Text & "Minimum = " & .Min(mdY) & vbCrLf
        'Maximum
        Text1.Text = Text1.Text & "Maximum = " & .Max(mdY) & vbCrLf
        'Median
        Text1.Text = Text1.Text & "Median = " & .Median(mdY) & vbCrLf
        'Average
        Text1.Text = Text1.Text & "Average = " & .Average(mdY)
    End With
End Sub


Private Sub Form_Load()
    'load X & Y arrays
    mdX(1) = 1
    mdX(2) = 2
    mdX(3) = 3
    mdX(4) = 4
    mdX(5) = 5
    mdX(6) = 6
    
    mdY(1) = 27
    mdY(2) = 34
    mdY(3) = 30
    mdY(4) = 33
    mdY(5) = 40
    mdY(6) = 52

    On Error Resume Next
    
    'crank-up Excel
    Set mobjExcel = GetObject(, "Excel.Application")
    
    If Err.Number Then
        Err.Clear
        Set mobjExcel = CreateObject("Excel.Application")
        If Err.Number Then
            MsgBox "Can't open Excel."
            Exit Sub
        End If
    End If
    
    'make invisible. Nobody has to know!
    mobjExcel.Visible = False
End Sub

Sub SetGraph()
Dim vGrid(1 To 6, 1) As Double
Dim ix As Integer
    
    'load vGrid array with line data
    For ix = 1 To 6
        'actual data points
        vGrid(ix, 0) = mdY(ix)
        'calculated regression line data
        vGrid(ix, 1) = mdY2(ix)
    Next ix
    
    With MSChart1
        .ChartType = 3 '2D line chart
        .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = mobjExcel.WorksheetFunction.Min(mdY) - 5
        .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = mobjExcel.WorksheetFunction.Max(mdY) + 10
        .Plot.Axis(VtChAxisIdY).CategoryScale.Auto = False
        .Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = 10
        .Title = Me.Caption
        .ChartData = vGrid()
        .ShowLegend = True
        .ColumnLabelCount = 2
        .Column = 1
        .ColumnLabel = "Data"
        .Column = 2
        .ColumnLabel = "RegLine"
        .Refresh
        .Visible = True
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'make sure we shut down Excel when leaving
    Set mobjExcel = Nothing
End Sub


