VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAveryLabels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Label printing class"
Option Explicit

Private m_Inprogress As Boolean     'Has a label been setup and not yet printed?
Private m_iBigFont As Integer       'First Row Font Size
Private m_iLittleFont As Integer    'Additional Rows Font Size

Private m_MaxCol As Integer         'Maximum # of columns of labels per page
Private m_MaxRow As Integer         'Maximum # of rows of labels per page
Private m_Top As Single             'Top margin (dist from top of page to 1st label)
Private m_LeftOne As Single         'Left margin (dist from left of page to 1st label)
Private m_HPitch As Single          'Pitch=Width of label + Space between label columns
Private m_VPitch As Single          'Pitch= Height of label + Space between labels vertically
Private m_LabelHeight As Single     'Height of label in inches
Private m_LabelWidth As Single      'Width of one label
Private Const TwipsPerInch = 1440   'Number of twips in one inch

Dim HP_Top_Border_Adjustment As Long    'top of page offset fudge factor
Property Get TwipsHigh() As Long             'gets top margin (in inches)
Attribute TwipsHigh.VB_Description = "Height of the label in twips."
    TwipsHigh = TwipsPerInch * m_LabelHeight
End Property
Property Get TwipsWide() As Long             'gets top margin (in inches)
Attribute TwipsWide.VB_Description = "Width of label in twips"
    TwipsWide = TwipsPerInch * m_LabelWidth
End Property
Property Get AvailablePointsHigh() As Long             'gets top margin (in inches)
Attribute AvailablePointsHigh.VB_Description = "Hight of label in printer's points."
    AvailablePointsHigh = (TwipsPerInch * m_LabelHeight) / 20
End Property
Property Get AvailablePointsWide() As Long             'gets num
Attribute AvailablePointsWide.VB_Description = "Width of label in printer's points"
    AvailablePointsWide = (TwipsPerInch * m_LabelWidth) / 20
End Property
Property Get AvailableLinesPerLabel() As Long             'gets max number of lines per label
Attribute AvailableLinesPerLabel.VB_Description = "Number of lines that will fit on a lable, given the bigFont and LittleFont attributes"
    AvailableLinesPerLabel = ((((TwipsPerInch * m_LabelHeight) / 20) - m_iBigFont)) / m_iLittleFont
End Property
Property Get AvailableCharactersPerLabelLine() As Long             'gets max number of lines per label
Attribute AvailableCharactersPerLabelLine.VB_Description = "The number of characters that will fit on a line on the label if a non-proportional font is used and the font size is set to the LittleFont attribute."
    AvailableCharactersPerLabelLine = ((TwipsPerInch * m_LabelWidth) / 20) / m_iLittleFont
End Property
Property Get AvailableCharactersPerFirstLabelLine() As Long             'gets max number of lines per label
Attribute AvailableCharactersPerFirstLabelLine.VB_Description = "The number of characters that will fit on a first  line on the label if a non-proportional font is used and the font size is set to the BigFont attribute."
    AvailableCharactersPerFirstLabelLine = ((TwipsPerInch * m_LabelWidth) / 20) / m_iBigFont
End Property
Property Let TopMargin(margin As Single)             'Sets top margin (in inches)
Attribute TopMargin.VB_Description = "Distance from top of form to top of topmost label."
    m_Top = margin
End Property
Property Get TopMargin() As Single             'gets top margin (in inches)
    TopMargin = m_Top
End Property
Property Let LeftMargin(margin As Single)          'Sets left margin (in inches)
Attribute LeftMargin.VB_Description = "The distance from the left edge of the form to the left edge of the left-most label."
    m_LeftOne = margin
End Property
Property Get LeftMargin() As Single          'gets left margin (in inches)
    LeftMargin = m_LeftOne
End Property
Property Let LabelHeight(height As Single)         'Sets height of each label (in inches)
Attribute LabelHeight.VB_Description = "The label height, in inches."
    m_LabelHeight = height
End Property
Property Get LabelHeight() As Single         'gets height of each label (in inches)
    LabelHeight = m_LabelHeight
End Property
Property Let LabelWidth(width As Single)           'Sets width of each label (in inches)
Attribute LabelWidth.VB_Description = "Width of the label, in inches."
    m_LabelWidth = width
End Property
Property Get LabelWidth() As Single           'gets width of each label (in inches)
    LabelWidth = m_LabelWidth
End Property
Property Let NumberOfColumns(number As Integer)     'Sets number of columns per page
Attribute NumberOfColumns.VB_Description = "Number of columns across the page"
    m_MaxCol = number
End Property
Property Get NumberOfColumns() As Integer    'gets number of columns per page
    NumberOfColumns = m_MaxCol
End Property
Property Let NumberOfRows(number As Integer)        'Sets number of rows per page
Attribute NumberOfRows.VB_Description = "Number of labels down the page"
    m_MaxRow = number
End Property
Property Get NumberOfRows() As Integer        'gets number of rows per page
    NumberOfRows = m_MaxRow
End Property
Property Let HPitch(distance As Single)             'Sets horizontal pitch (in inches)
Attribute HPitch.VB_Description = "Horizontal pitch.  The distance between the left edges of adjacent labels on a page.  Includes the width of the label and the filler between labels.  HPitch is in inches"
    m_HPitch = distance
End Property
Property Get HPitch() As Single             'gets horizontal pitch (in inches)
    HPitch = m_HPitch
End Property
Property Let VPitch(distance As Single)             'Sets vertical pitch (in inches)
Attribute VPitch.VB_Description = "Vertical pitch of the label.. That is, the distance, in inches, from the top of a label to the top of the label below it on the page."
    m_VPitch = distance
End Property
Property Get VPitch() As Single            'gets vertical pitch (in inches)
    VPitch = m_VPitch
End Property
Property Let LittleFont(fsize As Integer)           'Sets fontsize of first label row
Attribute LittleFont.VB_Description = "Sets the fontsize (in points) of the second and succeeding label lines."
    m_iLittleFont = fsize
End Property
Property Get LittleFont() As Integer           'gets fontsize of first label row
    LittleFont = m_iLittleFont
End Property
Property Let BigFont(fsize As Integer)              'Sets fontsize of other label rows
Attribute BigFont.VB_Description = "The fontsize (in points) of the first (top) label line."
    m_iBigFont = fsize
End Property
Property Get BigFont() As Integer              'gets fontsize of other label rows
    BigFont = m_iBigFont
End Property
Property Get TopBorderAdjustment() As Long          'Returns top border fudge factor (twips)
Attribute TopBorderAdjustment.VB_Description = "Offset from top of page.  (Don't forget that it CAN be negative!)."
    TopBorderAdjustment = HP_Top_Border_Adjustment
End Property
Property Let TopBorderAdjustment(newValue As Long)  'Sets top border fudge factor (twips)
    HP_Top_Border_Adjustment = newValue
End Property
Property Get JobStarted() As Boolean                'Returns in-process flag
Attribute JobStarted.VB_Description = "Has a label been setup and not printed?"
    JobStarted = m_Inprogress
End Property
Sub PageFinished()                                  'Finish page and clear in-process flag
Attribute PageFinished.VB_Description = "Output the page to the printer."
    Printer.EndDoc
    m_Inprogress = False
End Sub
Sub LabelPrint(TargRow As Integer, TargCol As Integer, TargLine As Integer, MessageData As String)
Attribute LabelPrint.VB_Description = "Insert a line of text into a label for printing.  "
    'This subroutine puts the data at the right place to print.
    '  The PageFinished routine above does the actual enddoc.
    
    'Make sure that we are within the labels rows/columns limits.
    If TargRow > m_MaxRow _
       Or TargCol > m_MaxCol _
       Or TargLine > ((((TwipsPerInch * m_LabelHeight) / 20) - m_iBigFont)) / m_iLittleFont _
       Then Exit Sub
    
    Dim x As Printer
    Dim i As Integer
          
    'set the inprogress flag as active
    m_Inprogress = True
    
    'setup variables for label positions
    Dim LabelX As Double
    Dim LabelY As Double
    
    'calculate vertical label position based on the label row
    LabelY = (TargRow - 1) * (TwipsPerInch * m_VPitch)
    
    'add any fudge factor for top of form
    LabelY = LabelY + HP_Top_Border_Adjustment
    
    ' add space for lines on label above this line
    If TargLine > 1 Then LabelY = LabelY + m_iBigFont * 20
    If TargLine > 2 Then
       For i = 1 To TargLine - 2
          LabelY = LabelY + m_iLittleFont * 20
          Next i
       End If
       
    'calculate horizontal label position based on the label column
    LabelX = TwipsPerInch * (m_LeftOne + (TargCol - 1) * m_HPitch)
    
    Printer.CurrentX = LabelX
    Printer.CurrentY = LabelY
    
    ' set the font
    If TargLine = 1 Then Printer.FontSize = m_iBigFont Else Printer.FontSize = m_iLittleFont
   
    'Place data onto printer form
    Printer.Print MessageData
    'NOTE: Once we print, the CurrentY changes to the next line . . .
    
    Printer.CurrentY = LabelY + Printer.TextHeight(MessageData)
    
End Sub

Private Sub Class_Initialize()
    m_Inprogress = False            'No labels initialized yet
    m_iBigFont = 36                 'Default to 36 point first row font
    m_iLittleFont = 10              'Default to 10 point second row font
    
    'HP_Top_Border_Adjustment = -1 * (0.125 * 1440#)
    'HP_Top_Border_Adjustment = (0.5 * 1440#)
    'HP_Top_Border_Adjustment = (0.25 * 1440#)
    'HP_Top_Border_Adjustment = (0.125 * 1440#)
    
    HP_Top_Border_Adjustment = (0.2 * TwipsPerInch)
End Sub

Private Sub Class_Terminate()
    If m_Inprogress = True Then
        Printer.EndDoc              'If labels have been initialized, print them
        m_Inprogress = False        '  clear in process flag
    End If
End Sub
