VERSION 5.00
Begin VB.UserControl AnalogClock 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   2505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2505
   EditAtDesignTime=   -1  'True
   KeyPreview      =   -1  'True
   ScaleHeight     =   2505
   ScaleWidth      =   2505
   ToolboxBitmap   =   "AnalogClock.ctx":0000
   Begin VB.Timer tmrAutoUpdate 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   30
      Top             =   30
   End
   Begin VB.PictureBox picClock 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1965
      Left            =   0
      MouseIcon       =   "AnalogClock.ctx":0312
      ScaleHeight     =   131
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   165
      TabIndex        =   0
      Top             =   510
      Width           =   2475
   End
End
Attribute VB_Name = "AnalogClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum BoxStyles
    None = 0
    Inset = 1
    Raised = 2
End Enum

Private Type ClockDevType
    TitleShow As Boolean
    TitleFormat As String
    TitleHeight As Integer
    TitleWidth As Integer
    TitleBoxStyle As BoxStyles
    TitleBevelWidth As Integer
    TitleBackColor As Long 'OLE_COLOR
    TitleForeColor As Long 'OLE_COLOR
    TitleFont As StdFont

    ClockHeight As Integer
    ClockWidth As Integer
    ClockBoxStyle As BoxStyles
    ClockBevelWidth As Integer
    ClockBackColor As Long 'OLE_COLOR
    ClockForeColor As Long 'OLE_COLOR
    ClockFont As StdFont

    AMPMShow As Boolean
    AMPMWidth As Integer
    AMPMHeight As Integer
    AMPMBoxStyle As BoxStyles
    AMPMBevelWidth As Integer
    AMPMFont As StdFont
    HLBackColor As Long 'OLE_COLOR
    HLForeColor As Long 'OLE_COLOR

    CentreX As Integer
    CentreY As Integer
    Radius As Integer

    HandHourColor As Long 'OLE_COLOR
    HandMinuteColor As Long 'OLE_COLOR
    HandSecondColor As Long 'OLE_COLOR
    HandHourWidth As Integer
    HandMinuteWidth As Integer
    HandSecondWidth As Integer
    HandSecondShow As Boolean

    HourBoxStyle As BoxStyles
    HourBevelWidth As Integer
    HourColor As Long 'OLE_COLOR

    MinuteColor As Long 'OLE_COLOR
    SecondColor As Long 'OLE_COLOR

    AMText As String
    PMText As String

    Time As Date
    AllowRotate As Boolean
    RotateHand As Integer       ' 1-minute, 2-hours
    AutoUpdate As Boolean       ' When true act like a normal clock
End Type

Private ClockDev As ClockDevType    ' The clock settings (this way you only have to remember one variable name and use the dot for the rest)
Private mnMouseX As Single
Private mnMouseY As Single
Private mbRedraw As Boolean         ' Flag to disable redraw - usefull when updating lots of properties (speeds up when switching off, when done, switch back on again)
Private mbInitialised As Boolean

Private Const Pi As Double = 3.14159265358979   ' How accurate can you get...

Event Change()
Event ClockGotFocus()
Event ClockLostFocus()

Private Sub UserControl_Initialize()
    picClock.ScaleMode = 3
    picClock.AutoRedraw = True
    With ClockDev
        .Time = TimeValue("00:00:00")
        .AMText = "am"                  ' Just in case AM/PM is called
        .PMText = "pm"                  ' something else in another country.
    End With
    mbRedraw = False
End Sub

Private Sub UserControl_Terminate()
    tmrAutoUpdate.Enabled = False       ' Switch it off for sure.
End Sub

Private Sub UserControl_InitProperties()
    With ClockDev
        .TitleHeight = 23
        .ClockHeight = 0
        .ClockWidth = 0

        .TitleShow = True
        .TitleFormat = "h:nn am/pm"
        .TitleWidth = picClock.ScaleWidth
        .TitleBoxStyle = Inset
        .TitleBevelWidth = 2
        .TitleBackColor = vbWindowBackground
        .TitleForeColor = vbButtonText
        Set .TitleFont = UserControl.Font

        .ClockBoxStyle = Raised
        .ClockBevelWidth = 2
        .ClockBackColor = vbButtonFace
        .ClockForeColor = vbButtonText
        Set .ClockFont = UserControl.Font

        .HandHourColor = vbBlack
        .HandHourWidth = 3
        .HandMinuteColor = vbBlack
        .HandMinuteWidth = 2
        .HandSecondColor = vbGrayText
        .HandSecondWidth = 1
        .HandSecondShow = False

        .HLBackColor = vbHighlight      ' Highlite background
        .HLForeColor = vbHighlightText  ' Highlite foreground
        .AMPMShow = True
        .AMPMWidth = 30
        .AMPMHeight = 23
        .AMPMBoxStyle = Inset
        .AMPMBevelWidth = 2
        Set .AMPMFont = UserControl.Font

        .HourBoxStyle = Inset
        .HourBevelWidth = 2
        .HourColor = vbButtonShadow

        .MinuteColor = vbButtonShadow
        .SecondColor = vbButtonShadow

        .Time = TimeValue("00:00:00")
        .AllowRotate = False
        .RotateHand = 0
    End With

    picClock.MousePointer = vbDefault

    mbInitialised = True
End Sub

Private Sub UserControl_Resize()
    picClock.Move 0, 0, UserControl.Width, UserControl.Height
    PaintClock
End Sub

Private Sub UserControl_Show()
    mbRedraw = True
    PaintClock
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        ClockDev.TitleHeight = .ReadProperty("TitleHeight", 23)
        ClockDev.TitleShow = .ReadProperty("TitleShow", True)
        ClockDev.TitleFormat = .ReadProperty("TitleFormat", "h:nn am/pm")
        ClockDev.TitleWidth = .ReadProperty("TitleWidth", picClock.ScaleWidth)
        ClockDev.TitleBoxStyle = .ReadProperty("TitleBoxStyle", Inset)
        ClockDev.TitleBevelWidth = .ReadProperty("TitleBevelWidth", 2)
        ClockDev.TitleBackColor = .ReadProperty("TitleBackColor", vbWindowBackground)
        ClockDev.TitleForeColor = .ReadProperty("TitleForeColor", vbButtonText)
        Set ClockDev.TitleFont = .ReadProperty("TitleFont", UserControl.Font)

        ClockDev.ClockBoxStyle = .ReadProperty("ClockBoxStyle", Raised)
        ClockDev.ClockBevelWidth = .ReadProperty("ClockBevelWidth", 2)
        ClockDev.ClockBackColor = .ReadProperty("ClockBackColor", vbButtonFace)
        ClockDev.ClockForeColor = .ReadProperty("ClockForeColor", vbButtonText)
        Set ClockDev.ClockFont = .ReadProperty("ClockFont", UserControl.Font)

        ClockDev.HandHourColor = .ReadProperty("HandHourColor", vbBlack)
        ClockDev.HandHourWidth = .ReadProperty("HandHourWidth", 3)
        ClockDev.HandMinuteColor = .ReadProperty("HandMinuteColor", vbBlack)
        ClockDev.HandMinuteWidth = .ReadProperty("HandMinuteWidth", 2)
        ClockDev.HandSecondColor = .ReadProperty("HandSecondColor", vbGrayText)
        ClockDev.HandSecondWidth = .ReadProperty("HandSecondWidth", 1)
        ClockDev.HandSecondShow = .ReadProperty("HandSecondShow", False)

        ClockDev.HLBackColor = .ReadProperty("HLBackColor", vbHighlight)
        ClockDev.HLForeColor = .ReadProperty("HLForeColor", vbHighlightText)
        ClockDev.AMPMShow = .ReadProperty("AMPMShow", True)
        ClockDev.AMPMWidth = .ReadProperty("AMPMWidth", 30)
        ClockDev.AMPMHeight = .ReadProperty("AMPMHeight", 23)
        ClockDev.AMPMBoxStyle = .ReadProperty("AMPMBoxStyle", Inset)
        ClockDev.AMPMBevelWidth = .ReadProperty("AMPMBevelWidth", 2)
        Set ClockDev.AMPMFont = .ReadProperty("AMPMFont", UserControl.Font)

        ClockDev.HourBoxStyle = .ReadProperty("HourBoxStyle", Inset)
        ClockDev.HourBevelWidth = .ReadProperty("HourBevelWidth", 2)
        ClockDev.HourColor = .ReadProperty("HourColor", vbButtonShadow)

        ClockDev.MinuteColor = .ReadProperty("MinuteColor", vbButtonShadow)
        ClockDev.SecondColor = .ReadProperty("SecondColor", vbButtonShadow)

        ClockDev.AMText = .ReadProperty("AMText", "am")
        ClockDev.PMText = .ReadProperty("PMText", "pm")

        ClockDev.AutoUpdate = .ReadProperty("AutoUpdate", False)
    End With

    If Ambient.UserMode Then
        tmrAutoUpdate.Enabled = ClockDev.AutoUpdate
        If tmrAutoUpdate.Enabled Then tmrAutoUpdate_Timer
    End If

    mbInitialised = True
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "TitleHeight", ClockDev.TitleHeight, 23
        .WriteProperty "TitleShow", ClockDev.TitleShow, True
        .WriteProperty "TitleFormat", ClockDev.TitleFormat, "h:nn am/pm"
        .WriteProperty "TitleWidth", ClockDev.TitleWidth, picClock.ScaleWidth
        .WriteProperty "TitleBoxStyle", ClockDev.TitleBoxStyle, Inset
        .WriteProperty "TitleBevelWidth", ClockDev.TitleBevelWidth, 2
        .WriteProperty "TitleBackColor", ClockDev.TitleBackColor, vbWindowBackground
        .WriteProperty "TitleForeColor", ClockDev.TitleForeColor, vbButtonText
        .WriteProperty "TitleFont", ClockDev.TitleFont, UserControl.Font

        .WriteProperty "ClockBoxStyle", ClockDev.ClockBoxStyle, Raised
        .WriteProperty "ClockBevelWidth", ClockDev.ClockBevelWidth, 2
        .WriteProperty "ClockBackColor", ClockDev.ClockBackColor, vbButtonFace
        .WriteProperty "ClockForeColor", ClockDev.ClockForeColor, vbButtonText
        .WriteProperty "ClockFont", ClockDev.ClockFont, UserControl.Font

        .WriteProperty "HandHourColor", ClockDev.HandHourColor, vbBlack
        .WriteProperty "HandHourWidth", ClockDev.HandHourWidth, 3
        .WriteProperty "HandMinuteColor", ClockDev.HandMinuteColor, vbBlack
        .WriteProperty "HandMinuteWidth", ClockDev.HandMinuteWidth, 2
        .WriteProperty "HandSecondColor", ClockDev.HandSecondColor, vbGrayText
        .WriteProperty "HandSecondWidth", ClockDev.HandSecondWidth, 1
        .WriteProperty "HandSecondShow", ClockDev.HandSecondShow, False

        .WriteProperty "HLBackColor", ClockDev.HLBackColor, vbHighlight
        .WriteProperty "HLForeColor", ClockDev.HLForeColor, vbHighlightText
        .WriteProperty "AMPMShow", ClockDev.AMPMShow, True
        .WriteProperty "AMPMWidth", ClockDev.AMPMWidth, 30
        .WriteProperty "AMPMHeight", ClockDev.AMPMHeight, 23
        .WriteProperty "AMPMBoxStyle", ClockDev.AMPMBoxStyle, Inset
        .WriteProperty "AMPMBevelWidth", ClockDev.AMPMBevelWidth, 2
        .WriteProperty "AMPMFont", ClockDev.AMPMFont, UserControl.Font

        .WriteProperty "HourBoxStyle", ClockDev.HourBoxStyle, Inset
        .WriteProperty "HourBevelWidth", ClockDev.HourBevelWidth, 2
        .WriteProperty "HourColor", ClockDev.HourColor, vbButtonShadow

        .WriteProperty "MinuteColor", ClockDev.MinuteColor, vbButtonShadow
        .WriteProperty "SecondColor", ClockDev.SecondColor, vbButtonShadow

        .WriteProperty "AMText", ClockDev.AMText, "am"
        .WriteProperty "PMText", ClockDev.PMText, "pm"

        .WriteProperty "AutoUpdate", ClockDev.AutoUpdate, False
    End With
End Sub

' -------------------------------------------------------------

' Automatic update of the time (using system time)
' Timer fires each second
Private Sub tmrAutoUpdate_Timer()
    If Not ClockDev.HandSecondShow Then
        ' Only change if minutes change
        If VBA.Format(Now, "hh:nn") = VBA.Format(ClockDev.Time, "hh:nn") Then Exit Sub
    End If
    ClockDev.Time = TimeValue(Now)
    PaintAMPM
    PaintHands
    If mbRedraw Then picClock.Refresh
End Sub

Private Sub picClock_GotFocus()
    RaiseEvent ClockGotFocus
End Sub

Private Sub picClock_LostFocus()
    RaiseEvent ClockLostFocus
End Sub

Private Sub picClock_DblClick()
    If tmrAutoUpdate.Enabled Then Exit Sub
    ClockMouseDown mnMouseX, mnMouseY, vbLeftButton ' Double click only happens with left mouse button
End Sub

Private Sub picClock_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mnMouseX = X: mnMouseY = Y
    'If tmrAutoUpdate.Enabled Then Exit Sub
    ClockMouseDown X, Y, Button
End Sub

Private Sub picClock_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmrAutoUpdate.Enabled Then
        picClock.MousePointer = vbDefault
        Exit Sub
    End If

    Dim nLeft As Long, nTop As Long, nRight As Long, nBottom As Long

    If ClockDev.TitleShow Then
        nLeft = ClockDev.ClockBevelWidth
        nTop = ClockDev.TitleHeight + ClockDev.ClockBevelWidth + 1
        nRight = Int(picClock.ScaleWidth) - ClockDev.ClockBevelWidth - 1
        nBottom = Int(picClock.ScaleHeight) - ClockDev.ClockBevelWidth - 1
    Else
        nLeft = ClockDev.ClockBevelWidth
        nTop = ClockDev.ClockBevelWidth
        nRight = Int(picClock.ScaleWidth) - ClockDev.ClockBevelWidth - 1
        nBottom = Int(picClock.ScaleHeight) - ClockDev.ClockBevelWidth - 1
    End If

    If X < nLeft Or X > nRight Then
        picClock.MousePointer = vbDefault
    ElseIf Y < nTop Or Y > nBottom Then
        picClock.MousePointer = vbDefault
    Else
        picClock.MousePointer = vbCustom
    End If

    If ClockDev.AllowRotate Then
        MoveHands X, Y, Button
        If mbRedraw Then picClock.Refresh
    End If
End Sub

Private Sub picClock_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClockDev.RotateHand = 0
    If ClockDev.AllowRotate And Not tmrAutoUpdate.Enabled Then
        MoveHands X, Y, Button
        If mbRedraw Then picClock.Refresh
    End If
    ClockDev.AllowRotate = False
End Sub

' -------------------------------------------------------------

Private Sub PaintClock()
    If Not mbInitialised Then Exit Sub

    Dim nX As Single
    Dim nY As Single
    Dim i As Integer
    Dim nDigit As Integer

    ' Calculate centre and radius
    '
    ClockDev.CentreX = picClock.ScaleWidth / 2

    If ClockDev.TitleShow Then
        ClockDev.CentreY = ((picClock.ScaleHeight - ClockDev.TitleHeight) / 2) + ClockDev.TitleHeight
        If ClockDev.CentreX < (picClock.ScaleHeight - ClockDev.TitleHeight) / 2 Then
            ClockDev.Radius = ClockDev.CentreX
        Else
            ClockDev.Radius = ClockDev.CentreY - ClockDev.TitleHeight
        End If
    Else
        ClockDev.CentreY = picClock.ScaleHeight / 2
        If ClockDev.CentreX < picClock.ScaleHeight / 2 Then
            ClockDev.Radius = ClockDev.CentreX
        Else
            ClockDev.Radius = ClockDev.CentreY
        End If
    End If
    ' Reduce radius with 28%
    ClockDev.Radius = ClockDev.Radius - (ClockDev.Radius * 0.28)

    If Not mbRedraw Then Exit Sub

    ' Paint the title and clock box
    ' (title will be shown only in the top)
    '
    If ClockDev.TitleShow Then
        PaintBox 0, 0, picClock.ScaleWidth - 1, ClockDev.TitleHeight - 1, ClockDev.TitleBackColor, ClockDev.TitleBoxStyle, ClockDev.TitleBevelWidth
        PaintBox 0, ClockDev.TitleHeight, Int(picClock.ScaleWidth) - 1, Int(picClock.ScaleHeight) - 1, ClockDev.ClockBackColor, ClockDev.ClockBoxStyle, ClockDev.ClockBevelWidth
    Else
        PaintBox 0, 0, Int(picClock.ScaleWidth) - 1, Int(picClock.ScaleHeight) - 1, ClockDev.ClockBackColor, ClockDev.ClockBoxStyle, ClockDev.ClockBevelWidth
    End If

    ' Paint minute dots
    '
    For i = 0 To 354 Step 6
        nX = Cos(i * (Pi / 180))
        nY = Sin(i * (Pi / 180))
        picClock.PSet (ClockDev.CentreX + (nX * ClockDev.Radius), (ClockDev.CentreY + (nY * ClockDev.Radius))), ClockDev.MinuteColor
    Next i
    
    ' Show the hour boxes and digits
    '
    SetFont ClockDev.ClockFont, ClockDev.ClockForeColor
    For i = -330 To 0 Step 30
        nX = Cos(i * (Pi / 180))
        nY = Sin(i * (Pi / 180))
        PaintBox ClockDev.CentreX + (nX * ClockDev.Radius) - 3, ClockDev.CentreY - (nY * ClockDev.Radius) - 3, ClockDev.CentreX + (nX * ClockDev.Radius) + 3, ClockDev.CentreY - (nY * ClockDev.Radius) + 3, ClockDev.HourColor, ClockDev.HourBoxStyle, ClockDev.HourBevelWidth
        Select Case i
        Case -330
            nDigit = 2
        Case -300
            nDigit = 1
        Case Else
            nDigit = Abs((i / 30) - 3)
        End Select
        PrintText ClockDev.CentreX + (nX * (ClockDev.Radius + 12)) - 3, ClockDev.CentreY - (nY * (ClockDev.Radius + 12)), ClockDev.CentreX + (nX * (ClockDev.Radius + 12)), ClockDev.CentreY - (nY * (ClockDev.Radius + 12)), Str(nDigit), "center", 0.5
    Next i

    ' Paint the AM/PM boxes
    '
    PaintAMPM

    ' Show the time (show hands) [also generates a Change event]
    '
    PaintHands False

    picClock.Refresh
End Sub

Private Sub PaintHands(Optional bNotify As Boolean = True)
    If Not mbRedraw Then
        If bNotify Then RaiseEvent Change
        Exit Sub
    End If

    Dim nHourAngle As Single
    Dim nMinuteAngle As Single
    Dim nHour As Single
    Dim nMinute As Integer
    Dim nEndX As Single
    Dim nEndY As Single

    On Local Error Resume Next

    ' Clear face
    picClock.FillColor = ClockDev.ClockBackColor
    picClock.Circle (ClockDev.CentreX, ClockDev.CentreY), ClockDev.Radius - 5, ClockDev.ClockBackColor

    If ClockDev.HandSecondShow Then
        ' Draw second hand
        picClock.DrawWidth = ClockDev.HandSecondWidth
        nMinute = VBA.Second(ClockDev.Time)
        Select Case nMinute
        Case 1 To 15
            nMinuteAngle = (15 - nMinute) * 6
        Case Else
            nMinuteAngle = Abs((nMinute - 75)) * 6
        End Select
        nEndX = Cos(nMinuteAngle * (Pi / 180))
        nEndY = Sin(nMinuteAngle * (Pi / 180))
        picClock.Line (ClockDev.CentreX, ClockDev.CentreY)-(ClockDev.CentreX + (nEndX * (ClockDev.Radius - 8)), ClockDev.CentreY - (nEndY * (ClockDev.Radius - 8))), ClockDev.HandSecondColor
    End If

    ' Draw minute hand
    picClock.DrawWidth = ClockDev.HandMinuteWidth
    nMinute = VBA.Minute(ClockDev.Time)
    Select Case nMinute
    Case 1 To 15
        nMinuteAngle = (15 - nMinute) * 6
    Case Else
        nMinuteAngle = Abs((nMinute - 75)) * 6
    End Select
    nEndX = Cos(nMinuteAngle * (Pi / 180))
    nEndY = Sin(nMinuteAngle * (Pi / 180))
    picClock.Line (ClockDev.CentreX, ClockDev.CentreY)-(ClockDev.CentreX + (nEndX * (ClockDev.Radius - 8)), ClockDev.CentreY - (nEndY * (ClockDev.Radius - 8))), IIf(ClockDev.RotateHand = 1, ClockDev.HLForeColor, ClockDev.HandMinuteColor)

    ' Draw hour hand
    picClock.DrawWidth = ClockDev.HandHourWidth
    nHour = VBA.Hour(ClockDev.Time)
    If nHour > 12 Then nHour = nHour - 12
    Select Case nHour
    Case 1, 2, 3
        nHour = nHour + (nMinute / 60)
        nHourAngle = (3 - nHour) * 30
    Case Else
        nHour = nHour + (nMinute / 60)
        nHourAngle = Abs((nHour - 15)) * 30
    End Select
    nEndX = Cos(nHourAngle * (Pi / 180))
    nEndY = Sin(nHourAngle * (Pi / 180))
    picClock.Line (ClockDev.CentreX, ClockDev.CentreY)-(ClockDev.CentreX + (nEndX * (ClockDev.Radius - 21)), ClockDev.CentreY - (nEndY * (ClockDev.Radius - 21))), IIf(ClockDev.RotateHand = 2, ClockDev.HLForeColor, ClockDev.HandHourColor)

    ' Little box in the centre of the clock (where the hands are attached)
    PaintBox ClockDev.CentreX - 3, ClockDev.CentreY - 3, ClockDev.CentreX + 3, ClockDev.CentreY + 3, ClockDev.HourColor, Raised, 2

    If ClockDev.TitleShow Then
        ' Place the text
        SetFont ClockDev.TitleFont, ClockDev.TitleForeColor
        PaintBox ClockDev.TitleBevelWidth + 2, ClockDev.TitleBevelWidth + 2, picClock.ScaleWidth - ClockDev.TitleBevelWidth - 2, ClockDev.TitleHeight - ClockDev.TitleBevelWidth - 2, ClockDev.TitleBackColor, None, 0
        PrintText ClockDev.TitleBevelWidth + 2, ClockDev.TitleBevelWidth + 2, picClock.ScaleWidth - ClockDev.TitleBevelWidth - 2, ClockDev.TitleHeight - ClockDev.TitleBevelWidth - 2, Format(ClockDev.Time, ClockDev.TitleFormat), "center", 0.5
    End If

    If bNotify Then RaiseEvent Change
End Sub

Private Sub PaintAMPM()
    If Not ClockDev.AMPMShow Or Not mbRedraw Then Exit Sub

    Dim nStartY As Long, nEndY As Long
    Dim nLeftX As Long, nRightX As Long

    ' Calculate box positions
    If ClockDev.ClockBevelWidth = 0 Then
        nStartY = IIf(ClockDev.TitleShow, ClockDev.TitleHeight, 0)
        nLeftX = 0
        nRightX = picClock.ScaleWidth - (ClockDev.AMPMWidth + 1)
    Else
        nStartY = IIf(ClockDev.TitleShow, ClockDev.TitleHeight, 0) + ClockDev.ClockBevelWidth + 1
        nLeftX = ClockDev.ClockBevelWidth + 1
        nRightX = picClock.ScaleWidth - (ClockDev.AMPMWidth + ClockDev.ClockBevelWidth + 2)
    End If
    nEndY = nStartY + ClockDev.AMPMHeight

    SetFont ClockDev.AMPMFont, ClockDev.ClockForeColor

    If VBA.Hour(ClockDev.Time) < 12 Then
        ' active: AM
        PaintBox nLeftX, nStartY, nLeftX + ClockDev.AMPMWidth, nEndY, ClockDev.HLBackColor, ClockDev.AMPMBoxStyle, ClockDev.AMPMBevelWidth
        PrintText nLeftX, nStartY - 1, nLeftX + ClockDev.AMPMWidth, nEndY, ClockDev.AMText, "center", 0.5, ClockDev.HLForeColor
        PaintBox nRightX, nStartY, nRightX + ClockDev.AMPMWidth, nEndY, ClockDev.ClockBackColor, ClockDev.AMPMBoxStyle, ClockDev.AMPMBevelWidth
        PrintText nRightX, nStartY - 1, nRightX + ClockDev.AMPMWidth, nEndY, ClockDev.PMText, "center", 0.5, ClockDev.ClockForeColor
    Else
        ' active: PM
        PaintBox nLeftX, nStartY, nLeftX + ClockDev.AMPMWidth, nEndY, ClockDev.ClockBackColor, ClockDev.AMPMBoxStyle, ClockDev.AMPMBevelWidth
        PrintText nLeftX, nStartY - 1, nLeftX + ClockDev.AMPMWidth, nEndY, ClockDev.AMText, "center", 0.5, ClockDev.ClockForeColor
        PaintBox nRightX, nStartY, nRightX + ClockDev.AMPMWidth, nEndY, ClockDev.HLBackColor, ClockDev.AMPMBoxStyle, ClockDev.AMPMBevelWidth
        PrintText nRightX, nStartY - 1, nRightX + ClockDev.AMPMWidth, nEndY, ClockDev.PMText, "center", 0.5, ClockDev.HLForeColor
    End If
End Sub

' ---------------------------------------------------------------
' Internal generic painting routines

Private Sub SetFont(FontData As StdFont, nForeColor As Long)
    On Local Error Resume Next
    Set picClock.Font = FontData
    picClock.ForeColor = nForeColor
End Sub

Private Sub PaintBox(ByVal nStartX As Integer, ByVal nStartY As Integer, ByVal nEndX As Integer, ByVal nEndY As Integer, ByVal nColor As Long, ByVal nStyle As BoxStyles, ByVal nBevelWidth As Integer)
    Dim i As Integer, nBevel As Integer

    picClock.FillColor = nColor
    picClock.FillStyle = 0
    picClock.DrawStyle = 0
    picClock.DrawWidth = 1
    picClock.Line (nStartX, nStartY)-(nEndX, nEndY), nColor, BF

    If nBevelWidth < 1 Or nStyle = None Then Exit Sub

    nBevel = nBevelWidth - 1

    Select Case nStyle  ' Paint bevel
    Case Raised
        ' Enhance the outside
        picClock.Line (nStartX, nStartY)-(nEndX - 1, nStartY), vb3DLight
        picClock.Line (nStartX, nStartY)-(nStartX, nEndY - 1), vb3DLight
        picClock.Line (nStartX, nEndY)-(nEndX + 1, nEndY), vb3DDKShadow ' vbBlack
        picClock.Line (nEndX, nStartY)-(nEndX, nEndY), vb3DDKShadow ' vbBlack

        ' Paint the shadow
        For i = 1 To nBevel
            picClock.Line (nStartX + i, nStartY + i)-(nEndX - i, nStartY + i), vb3DHighlight ' RGB(255, 255, 255)
            picClock.Line (nStartX + i, nStartY + i)-(nStartX + i, nEndY - i), vb3DHighlight ' RGB(255, 255, 255)
            picClock.Line (nStartX + i, nEndY - i)-(nEndX - i + 1, nEndY - i), vbButtonShadow ' RGB(92, 92, 92)
            picClock.Line (nEndX - i, nStartY + i)-(nEndX - i, nEndY - i), vbButtonShadow ' RGB(92, 92, 92)
        Next i

    Case Inset
        ' Paint the shadow
        For i = 0 To (nBevel - 1)
            picClock.Line (nStartX + i, nStartY + i)-(nEndX - i, nStartY + i), vbButtonShadow
            picClock.Line (nStartX + i, nStartY + i)-(nStartX + i, nEndY - i), vbButtonShadow
            picClock.Line (nStartX + i, nEndY - i)-(nEndX - i + 1, nEndY - i), vb3DHighlight
            picClock.Line (nEndX - i, nStartY + i)-(nEndX - i, nEndY - i + 1), vb3DHighlight
        Next i

        ' Enhance the inside
        picClock.Line (nStartX + nBevel, nStartY + nBevel)-(nEndX - nBevel, nStartY + nBevel), vb3DDKShadow
        picClock.Line (nStartX + nBevel, nStartY + nBevel)-(nStartX + nBevel, nEndY - nBevel), vb3DDKShadow
        picClock.Line (nStartX + nBevel, nEndY - nBevel)-(nEndX - nBevel + 1, nEndY - nBevel), vb3DLight
        picClock.Line (nEndX - nBevel, nStartY + nBevel)-(nEndX - nBevel, nEndY - nBevel + 1), vb3DLight
    End Select
End Sub

Private Sub PrintText(ByVal nXpos As Integer, ByVal nYpos As Integer, ByVal nEndX As Integer, ByVal nEndY As Integer, ByVal sText As String, ByVal sAlignment As String, ByVal nVerticalFactor As Single, Optional ByVal nColor)
    Select Case LCase(sAlignment)
    Case "left"
        picClock.CurrentX = nXpos
        picClock.CurrentY = nYpos + (((nEndY - nYpos) * nVerticalFactor) - (picClock.TextHeight(sText) * nVerticalFactor))
    Case "right"
        picClock.CurrentX = nEndX - picClock.TextWidth(sText)
        picClock.CurrentY = nYpos + (((nEndY - nYpos) * nVerticalFactor) - (picClock.TextHeight(sText) * nVerticalFactor))
    Case "center"
        picClock.CurrentX = nXpos + (((nEndX - nXpos) / 2) - (picClock.TextWidth(sText) / 2))
        picClock.CurrentY = nYpos + (((nEndY - nYpos) * nVerticalFactor) - (picClock.TextHeight(sText) * nVerticalFactor))
    End Select
    If Not IsMissing(nColor) Then picClock.ForeColor = nColor   ' Color override...
    picClock.Print sText
End Sub

' ---------------------------------------------------------------

Private Sub ClockMouseDown(nXpos As Single, nYpos As Single, nButton As Integer)
    If tmrAutoUpdate.Enabled Then
        ClockDev.AllowRotate = False
        ClockDev.RotateHand = 0
        Exit Sub
    End If
    ClockDev.RotateHand = 0

    If ClockDev.TitleShow Then
        If nYpos > ClockDev.TitleHeight + ClockDev.ClockBevelWidth + 1 And nYpos < picClock.ScaleHeight - ClockDev.ClockBevelWidth - 1 And nXpos > ClockDev.ClockBevelWidth + 1 And nXpos < picClock.ScaleWidth - ClockDev.ClockBevelWidth - 1 Then
            If ClockDev.AMPMShow And nYpos < ClockDev.TitleHeight + ClockDev.AMPMHeight Then
                If (nXpos > ClockDev.ClockBevelWidth + 2 And nXpos < picClock.ScaleLeft + ClockDev.AMPMWidth + ClockDev.ClockBevelWidth) Then
                    ClockDev.AllowRotate = False
                    If VBA.Hour(ClockDev.Time) > 11 Then
                        ' Time is PM: change to AM
                        ClockDev.Time = ClockDev.Time - TimeSerial(12, 0, 0)
                        PaintAMPM
                        PaintHands
                        If mbRedraw Then picClock.Refresh
                    End If
                ElseIf (nXpos > picClock.ScaleWidth - ClockDev.AMPMWidth - ClockDev.ClockBevelWidth - 2 And nXpos < picClock.ScaleWidth - ClockDev.ClockBevelWidth - 4) Then
                    ClockDev.AllowRotate = False
                    If VBA.Hour(ClockDev.Time) < 12 Then
                        ' Time is AM: change to PM
                        ClockDev.Time = ClockDev.Time + TimeSerial(12, 0, 0)
                        PaintAMPM
                        PaintHands
                        If mbRedraw Then picClock.Refresh
                    End If
                Else
                    ClockDev.AllowRotate = True
                    ClockDev.RotateHand = nButton
                End If
            Else
                ClockDev.AllowRotate = True
                ClockDev.RotateHand = nButton
            End If
        End If
    Else
        If nYpos > ClockDev.ClockBevelWidth + 1 And nYpos < picClock.ScaleHeight - ClockDev.ClockBevelWidth - 1 And nXpos > ClockDev.ClockBevelWidth + 1 And nXpos < picClock.ScaleWidth - ClockDev.ClockBevelWidth - 1 Then
            If ClockDev.AMPMShow And nYpos < ClockDev.AMPMHeight Then
                If (nXpos > ClockDev.ClockBevelWidth + 2 And nXpos < picClock.ScaleLeft + ClockDev.AMPMWidth + ClockDev.ClockBevelWidth) Then
                    ClockDev.AllowRotate = False
                    If VBA.Hour(ClockDev.Time) > 11 Then
                        ' Time is PM: change to AM
                        ClockDev.Time = ClockDev.Time - TimeSerial(12, 0, 0)
                        PaintAMPM
                        PaintHands
                        If mbRedraw Then picClock.Refresh
                    End If
                ElseIf (nXpos > picClock.ScaleWidth - ClockDev.AMPMWidth - ClockDev.ClockBevelWidth - 2 And nXpos < picClock.ScaleWidth - ClockDev.ClockBevelWidth - 4) Then
                    ClockDev.AllowRotate = False
                    If VBA.Hour(ClockDev.Time) < 12 Then
                        ' Time is AM: change to PM
                        ClockDev.Time = ClockDev.Time + TimeSerial(12, 0, 0)
                        PaintAMPM
                        PaintHands
                        If mbRedraw Then picClock.Refresh
                    End If
                Else
                    ClockDev.AllowRotate = True
                    ClockDev.RotateHand = nButton
                End If
            Else
                ClockDev.AllowRotate = True
                ClockDev.RotateHand = nButton
            End If
        End If
    End If
End Sub

Private Sub MoveHands(nXpos As Single, nYpos As Single, nButton As Integer)
    If tmrAutoUpdate.Enabled Then Exit Sub
    If nButton < 1 Or nButton > 2 Then Exit Sub

    Dim nAngle As Integer
    Dim nX As Integer
    Dim nY As Integer
    Dim nMinutes As Integer
    Dim nHours As Integer

    nX = nXpos - ClockDev.CentreX
    nY = nYpos - ClockDev.CentreY

    If nX > 0 And nY < 0 Then
        nAngle = Atn(Abs(nY / nX)) * (180 / Pi)
    ElseIf nX < 0 And nY <= 0 Then
        nAngle = 180 - (Atn(Abs(nY / nX)) * (180 / Pi))
    ElseIf nX < 0 And nY > 0 Then
        nAngle = 180 + (Atn(Abs(nY / nX)) * (180 / Pi))
    ElseIf nX > 0 And nY > 0 Then
        nAngle = 360 - (Atn(Abs(nY / nX)) * (180 / Pi))
    ElseIf nX > 0 And nY = 0 Then
        nAngle = 0
    ElseIf nX = 0 And nY > 0 Then
        nAngle = 270
    ElseIf nX < 0 And nY = 0 Then
        nAngle = 180
    ElseIf nX = 0 And nY < 0 Then
        nAngle = 90
    End If

    Select Case nButton
    Case 1                  ' Adjust minutes
        If nAngle >= 0 And nAngle <= 90 Then
            nMinutes = Abs(nAngle - 90) / 6
        Else
            nMinutes = (450 - nAngle) / 6
        End If
        If nMinutes = 60 Then nMinutes = 0
        nHours = VBA.Hour(ClockDev.Time)

    Case 2                  ' Adjust hours
        nAngle = nAngle - 15
        If nAngle >= 0 And nAngle <= 90 Then
            nHours = Int(Abs(nAngle - 90) / 30)
        Else
            nHours = Int((450 - nAngle) / 30)
        End If
        If nHours > 12 Then nHours = nHours - 12
        If VBA.Hour(ClockDev.Time) > 11 Then nHours = nHours + 12
        nMinutes = VBA.Minute(ClockDev.Time)
    End Select

    ClockDev.Time = TimeSerial(nHours, nMinutes, 0)
    PaintHands
End Sub

' ---------------------------------------------------------------
' Properties....

Public Property Get Value() As Date
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Data"
Attribute Value.VB_UserMemId = 0
    On Local Error Resume Next
    Value = ClockDev.Time
End Property
Public Property Let Value(ByVal vNewValue As Date)
    ClockDev.Time = TimeValue(vNewValue)
    PaintAMPM
    PaintHands
    If mbRedraw Then picClock.Refresh
End Property

Public Property Get Time() As Date
    On Local Error Resume Next
    Time = ClockDev.Time
End Property
Public Property Let Time(ByVal vNewValue As Date)
    ClockDev.Time = TimeValue(vNewValue)
    PaintAMPM
    PaintHands False   ' Do NOT generate "Change" event
    If mbRedraw Then picClock.Refresh
End Property

Public Property Get TitleFormat() As String
Attribute TitleFormat.VB_ProcData.VB_Invoke_Property = ";Appearance"
    TitleFormat = ClockDev.TitleFormat
End Property
Public Property Let TitleFormat(ByVal vNewValue As String)
    ClockDev.TitleFormat = vNewValue
    PaintClock
    PropertyChanged "TitleFormat"
End Property

Public Property Get TitleShow() As Boolean
Attribute TitleShow.VB_ProcData.VB_Invoke_Property = ";Appearance"
    TitleShow = ClockDev.TitleShow
End Property
Public Property Let TitleShow(ByVal vNewValue As Boolean)
    ClockDev.TitleShow = vNewValue
    PaintClock
    PropertyChanged "TitleShow"
End Property

Public Property Get AMPMShow() As Boolean
Attribute AMPMShow.VB_ProcData.VB_Invoke_Property = ";Appearance"
    AMPMShow = ClockDev.AMPMShow
End Property
Public Property Let AMPMShow(ByVal vNewValue As Boolean)
    ClockDev.AMPMShow = vNewValue
    PaintClock
    PropertyChanged "AMPMShow"
End Property

Public Property Get HandSecondShow() As Boolean
Attribute HandSecondShow.VB_ProcData.VB_Invoke_Property = ";Appearance"
    HandSecondShow = ClockDev.HandSecondShow
End Property
Public Property Let HandSecondShow(ByVal vNewValue As Boolean)
    ClockDev.HandSecondShow = vNewValue
    PaintClock
    PropertyChanged "HandSecondShow"
End Property

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Public Property Get TitleBackColor() As OLE_COLOR
Attribute TitleBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    TitleBackColor = ClockDev.TitleBackColor
End Property
Public Property Let TitleBackColor(ByVal vNewValue As OLE_COLOR)
    ClockDev.TitleBackColor = vNewValue
    PaintClock
    PropertyChanged "TitleBackColor"
End Property

Public Property Get TitleForeColor() As OLE_COLOR
Attribute TitleForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    TitleForeColor = ClockDev.TitleForeColor
End Property
Public Property Let TitleForeColor(ByVal vNewValue As OLE_COLOR)
    ClockDev.TitleForeColor = vNewValue
    PaintClock
    PropertyChanged "TitleForeColor"
End Property

Public Property Get ClockBackColor() As OLE_COLOR
Attribute ClockBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ClockBackColor = ClockDev.ClockBackColor
End Property
Public Property Let ClockBackColor(ByVal vNewValue As OLE_COLOR)
    ClockDev.ClockBackColor = vNewValue
    PaintClock
    PropertyChanged "ClockBackColor"
End Property

Public Property Get ClockForeColor() As OLE_COLOR
Attribute ClockForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ClockForeColor = ClockDev.ClockForeColor
End Property
Public Property Let ClockForeColor(ByVal vNewValue As OLE_COLOR)
    ClockDev.ClockForeColor = vNewValue
    PaintClock
    PropertyChanged "ClockForeColor"
End Property

Public Property Get HandHourColor() As OLE_COLOR
Attribute HandHourColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    HandHourColor = ClockDev.HandHourColor
End Property
Public Property Let HandHourColor(ByVal vNewValue As OLE_COLOR)
    ClockDev.HandHourColor = vNewValue
    PaintHands
    If mbRedraw Then picClock.Refresh
    PropertyChanged "HandHourColor"
End Property

Public Property Get HandMinuteColor() As OLE_COLOR
Attribute HandMinuteColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    HandMinuteColor = ClockDev.HandMinuteColor
End Property
Public Property Let HandMinuteColor(ByVal vNewValue As OLE_COLOR)
    ClockDev.HandMinuteColor = vNewValue
    PaintHands
    If mbRedraw Then picClock.Refresh
    PropertyChanged "HandMinuteColor"
End Property

Public Property Get HandSecondColor() As OLE_COLOR
Attribute HandSecondColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    HandSecondColor = ClockDev.HandSecondColor
End Property
Public Property Let HandSecondColor(ByVal vNewValue As OLE_COLOR)
    ClockDev.HandSecondColor = vNewValue
    PaintHands
    If mbRedraw Then picClock.Refresh
    PropertyChanged "HandSecondColor"
End Property

Public Property Get HoursColor() As OLE_COLOR
Attribute HoursColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    HoursColor = ClockDev.HourColor
End Property
Public Property Let HoursColor(ByVal vNewValue As OLE_COLOR)
    ClockDev.HourColor = vNewValue
    PaintClock
    PropertyChanged "HourColor"
End Property

Public Property Get MinutesColor() As OLE_COLOR
Attribute MinutesColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    MinutesColor = ClockDev.MinuteColor
End Property
Public Property Let MinutesColor(ByVal vNewValue As OLE_COLOR)
    ClockDev.MinuteColor = vNewValue
    PaintClock
    PropertyChanged "MinuteColor"
End Property

Public Property Get HighliteBackColor() As OLE_COLOR
Attribute HighliteBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    HighliteBackColor = ClockDev.HLBackColor
End Property
Public Property Let HighliteBackColor(ByVal vNewValue As OLE_COLOR)
    ClockDev.HLBackColor = vNewValue
    PaintAMPM
    If mbRedraw Then picClock.Refresh
    PropertyChanged "HLBackColor"
End Property

Public Property Get HighliteForeColor() As OLE_COLOR
Attribute HighliteForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    HighliteForeColor = ClockDev.HLForeColor
End Property
Public Property Let HighliteForeColor(ByVal vNewValue As OLE_COLOR)
    ClockDev.HLForeColor = vNewValue
    PaintAMPM
    If mbRedraw Then picClock.Refresh
    PropertyChanged "HLForeColor"
End Property

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Public Property Get TitleBoxStyle() As BoxStyles
Attribute TitleBoxStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    TitleBoxStyle = ClockDev.TitleBoxStyle
End Property
Public Property Let TitleBoxStyle(ByVal vNewValue As BoxStyles)
    ClockDev.TitleBoxStyle = vNewValue
    PaintClock
    PropertyChanged "TitleBoxStyle"
End Property

Public Property Get TitleBevelWidth() As Integer
Attribute TitleBevelWidth.VB_ProcData.VB_Invoke_Property = ";Appearance"
    TitleBevelWidth = ClockDev.TitleBevelWidth
End Property
Public Property Let TitleBevelWidth(ByVal vNewValue As Integer)
    If vNewValue < 0 Then Exit Property
    ClockDev.TitleBevelWidth = vNewValue
    PaintClock
    PropertyChanged "TitleBevelWidth"
End Property

Public Property Get AMPMBoxStyle() As BoxStyles
Attribute AMPMBoxStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    AMPMBoxStyle = ClockDev.AMPMBoxStyle
End Property
Public Property Let AMPMBoxStyle(ByVal vNewValue As BoxStyles)
    ClockDev.AMPMBoxStyle = vNewValue
    PaintAMPM
    If mbRedraw Then picClock.Refresh
    PropertyChanged "AMPMBoxStyle"
End Property

Public Property Get AMPMBevelWidth() As Integer
Attribute AMPMBevelWidth.VB_ProcData.VB_Invoke_Property = ";Appearance"
    AMPMBevelWidth = ClockDev.AMPMBevelWidth
End Property
Public Property Let AMPMBevelWidth(ByVal vNewValue As Integer)
    If vNewValue < 0 Then Exit Property
    ClockDev.AMPMBevelWidth = vNewValue
    PaintAMPM
    If mbRedraw Then picClock.Refresh
    PropertyChanged "AMPMBevelWidth"
End Property

Public Property Get ClockBoxStyle() As BoxStyles
Attribute ClockBoxStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ClockBoxStyle = ClockDev.ClockBoxStyle
End Property
Public Property Let ClockBoxStyle(ByVal vNewValue As BoxStyles)
    ClockDev.ClockBoxStyle = vNewValue
    PaintClock
    PropertyChanged "ClockBoxStyle"
End Property

Public Property Get ClockBevelWidth() As Integer
Attribute ClockBevelWidth.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ClockBevelWidth = ClockDev.ClockBevelWidth
End Property
Public Property Let ClockBevelWidth(ByVal vNewValue As Integer)
    If vNewValue < 0 Then Exit Property
    ClockDev.ClockBevelWidth = vNewValue
    PaintClock
    PropertyChanged "ClockBevelWidth"
End Property

Public Property Get HoursBoxStyle() As BoxStyles
Attribute HoursBoxStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    HoursBoxStyle = ClockDev.HourBoxStyle
End Property
Public Property Let HoursBoxStyle(ByVal vNewValue As BoxStyles)
    ClockDev.HourBoxStyle = vNewValue
    PaintClock
    PropertyChanged "HourBoxStyle"
End Property

Public Property Get HoursBevelWidth() As Integer
Attribute HoursBevelWidth.VB_ProcData.VB_Invoke_Property = ";Appearance"
    HoursBevelWidth = ClockDev.HourBevelWidth
End Property
Public Property Let HoursBevelWidth(ByVal vNewValue As Integer)
    If vNewValue < 0 Then Exit Property
    ClockDev.HourBevelWidth = vNewValue
    PaintClock
    PropertyChanged "HourBevelWidth"
End Property

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Public Property Get HandHourWidth() As Integer
Attribute HandHourWidth.VB_ProcData.VB_Invoke_Property = ";Appearance"
    HandHourWidth = ClockDev.HandHourWidth
End Property
Public Property Let HandHourWidth(ByVal vNewValue As Integer)
    If vNewValue < 1 Then Exit Property
    ClockDev.HandHourWidth = vNewValue
    PaintHands
    If mbRedraw Then picClock.Refresh
    PropertyChanged "HandHourWidth"
End Property

Public Property Get HandMinuteWidth() As Integer
Attribute HandMinuteWidth.VB_ProcData.VB_Invoke_Property = ";Appearance"
    HandMinuteWidth = ClockDev.HandMinuteWidth
End Property
Public Property Let HandMinuteWidth(ByVal vNewValue As Integer)
    If vNewValue < 1 Then Exit Property
    ClockDev.HandMinuteWidth = vNewValue
    PaintHands
    If mbRedraw Then picClock.Refresh
    PropertyChanged "HandMinuteWidth"
End Property

Public Property Get HandSecondWidth() As Integer
Attribute HandSecondWidth.VB_ProcData.VB_Invoke_Property = ";Appearance"
    HandSecondWidth = ClockDev.HandSecondWidth
End Property
Public Property Let HandSecondWidth(ByVal vNewValue As Integer)
    If vNewValue < 1 Then Exit Property
    ClockDev.HandSecondWidth = vNewValue
    PaintHands
    If mbRedraw Then picClock.Refresh
    PropertyChanged "HandSecondWidth"
End Property

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Public Property Get TitleFont() As StdFont
Attribute TitleFont.VB_ProcData.VB_Invoke_Property = ";Font"
    Set TitleFont = ClockDev.TitleFont
End Property
Public Property Set TitleFont(ByVal vNewValue As StdFont)
    Set ClockDev.TitleFont = vNewValue
    PaintClock
    PropertyChanged "TitleFont"
End Property

Public Property Get AMPMFont() As StdFont
Attribute AMPMFont.VB_ProcData.VB_Invoke_Property = ";Font"
    Set AMPMFont = ClockDev.AMPMFont
End Property
Public Property Set AMPMFont(ByVal vNewValue As StdFont)
    Set ClockDev.AMPMFont = vNewValue
    PaintClock
    PropertyChanged "AMPMFont"
End Property

Public Property Get ClockFont() As StdFont
Attribute ClockFont.VB_ProcData.VB_Invoke_Property = ";Font"
    Set ClockFont = ClockDev.ClockFont
End Property
Public Property Set ClockFont(ByVal vNewValue As StdFont)
    Set ClockDev.ClockFont = vNewValue
    PaintClock
    PropertyChanged "ClockFont"
End Property

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Public Property Get AMPMWidth() As Integer
Attribute AMPMWidth.VB_ProcData.VB_Invoke_Property = ";Appearance"
    AMPMWidth = ClockDev.AMPMWidth * Screen.TwipsPerPixelX
End Property
Public Property Let AMPMWidth(ByVal vNewValue As Integer)
    vNewValue = vNewValue / Screen.TwipsPerPixelX
    If vNewValue < ((ClockDev.AMPMBevelWidth * 2) + 1) Then Exit Property
    ClockDev.AMPMWidth = vNewValue
    PaintClock
    PropertyChanged "AMPMWidth"
End Property

Public Property Get AMPMHeight() As Integer
Attribute AMPMHeight.VB_ProcData.VB_Invoke_Property = ";Appearance"
    AMPMHeight = ClockDev.AMPMHeight * Screen.TwipsPerPixelY
End Property
Public Property Let AMPMHeight(ByVal vNewValue As Integer)
    vNewValue = vNewValue / Screen.TwipsPerPixelY
    If vNewValue < ((ClockDev.AMPMBevelWidth * 2) + 1) Then Exit Property
    ClockDev.AMPMHeight = vNewValue
    PaintClock
    PropertyChanged "AMPMHeight"
End Property

Public Property Get TitleHeight() As Long
Attribute TitleHeight.VB_ProcData.VB_Invoke_Property = ";Appearance"
    TitleHeight = ClockDev.TitleHeight * Screen.TwipsPerPixelY
End Property
Public Property Let TitleHeight(ByVal vNewValue As Long)
    vNewValue = vNewValue / Screen.TwipsPerPixelY
    If vNewValue < ((ClockDev.TitleBevelWidth * 2) + 1) Then Exit Property
    ClockDev.TitleHeight = vNewValue
    PaintClock
    PropertyChanged "TitleHeight"
End Property

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Public Property Get Redraw() As Boolean
    Redraw = mbRedraw
End Property
Public Property Let Redraw(ByVal vNewValue As Boolean)
    mbRedraw = vNewValue
    If mbRedraw Then PaintClock
End Property

Public Property Get AutoUpdate() As Boolean
Attribute AutoUpdate.VB_ProcData.VB_Invoke_Property = ";Behavior"
    AutoUpdate = ClockDev.AutoUpdate
End Property
Public Property Let AutoUpdate(ByVal vNewValue As Boolean)
    ClockDev.AutoUpdate = vNewValue
    If Ambient.UserMode Then
        tmrAutoUpdate.Enabled = ClockDev.AutoUpdate
        If tmrAutoUpdate.Enabled Then tmrAutoUpdate_Timer
    End If
    PropertyChanged "AutoUpdate"
End Property

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' E.O.F
