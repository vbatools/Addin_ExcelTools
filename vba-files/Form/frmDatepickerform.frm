VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDatepickerform 
   ClientHeight    =   9210.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7545
   OleObjectBlob   =   "frmDatepickerform.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDatepickerform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Public DateGlobal   As Date

Public pickerMode   As Integer

Public colorWhite   As XlRgbColor

Public ColorMove    As XlRgbColor

Public ColorSelectedDate As XlRgbColor

Public ForeColorDayNo As XlRgbColor

Public ForeColorDayYes As XlRgbColor

Private minDate     As Date

Private maxDate     As Date

Private iDayForm    As Integer

Private iItem       As Integer

Private iMounthGlobal As Integer

Private iYearGlobal As Integer

Private sDateGlobal As String

Public iSelectGlobal As Integer

Private iShiftGlobal As Integer



Private Sub monthTitle_Click()

    If lbTypeCalendar.Caption = "1" Then Exit Sub

    iSelectGlobal = iMounthGlobal

    iItem = 0

    SetShowPicerMode1

End Sub



Private Sub yearTitle_Click()

    If lbTypeCalendar.Caption = "1" Then Exit Sub

    iSelectGlobal = iYearGlobal

    iItem = 0

    SetShowPicerMode2

End Sub



Private Sub nextMonthButton_Click()

    If pickerMode = 0 Then

        iMounthGlobal = iMounthGlobal + 1

        If iMounthGlobal > 12 Then

            iMounthGlobal = 1

            iYearGlobal = iYearGlobal + 1

        End If

        SetDayOneAndDayTwo

    ElseIf pickerMode = 2 Then

        iShiftGlobal = iShiftGlobal + 3

        Call SetYears

    End If

End Sub

Private Sub prevMonthButton_Click()

    If pickerMode = 0 Then

        iMounthGlobal = iMounthGlobal - 1

        If iMounthGlobal < 1 Then

            iMounthGlobal = 12

            iYearGlobal = iYearGlobal - 1

        End If

        SetDayOneAndDayTwo

    ElseIf pickerMode = 2 Then

        iShiftGlobal = iShiftGlobal - 3

        Call SetYears

    End If

End Sub



Private Sub UserForm_Initialize()

    colorWhite = 16777215

    ColorMove = 14737632

    ColorSelectedDate = 12632319

    ForeColorDayNo = 8421504

    ForeColorDayYes = -2147483630

End Sub



Private Sub UserForm_Activate()

    With Me

        .Width = redBG.Width

        .Height = barHeight.Height

    End With



    minDate = VBA.CDate(lbMinDate.Caption)

    maxDate = VBA.CDate(lbMaxDate.Caption)



    iMounthGlobal = VBA.Month(DateGlobal)

    iYearGlobal = VBA.Year(DateGlobal)



    SetCapitonControlSunday

    SetDayOneAndDayTwo

    todayButton.Caption = VBA.WeekdayName(VBA.Weekday(VBA.Date, vbMonday)) & ", " & VBA.day(VBA.Date) & " " & VBA.MonthName(VBA.Month(VBA.Date)) & " " & VBA.Year(VBA.Date)

    timeButton.Caption = VBA.Time

    Call SwitchVisibleButton(pickerMode)

    ' Центрирование формы
    Call CenterUserForm(Me)
End Sub



Private Sub SetControlBackColorUnderline(day)

    If iDayForm <> 0 Then

        If iDayForm <> day Then

            If iDayForm <= 67 Then

                Me.Controls("daybg" & iDayForm).BackColor = colorWhite

                SwitchColorControl Me.Controls("daybg" & iDayForm)

            ElseIf iDayForm = 68 Then

                datetimebg.BackColor = colorWhite

            ElseIf iDayForm = 69 Then

                monthTitle.Font.Underline = False

            ElseIf iDayForm = 70 Then

                yearTitle.Font.Underline = False

            End If

        End If

    End If

    If day > 0 Then

        If day <= 67 Then

            Me.Controls("daybg" & day).BackColor = ColorMove

        ElseIf day = 68 Then

            datetimebg.BackColor = ColorMove

        ElseIf day = 69 Then

            monthTitle.Font.Underline = True

        ElseIf day = 70 Then

            yearTitle.Font.Underline = True

        End If

        iDayForm = day

    End If

End Sub

Private Sub SetControlBackColor(itemIndex As Integer)

    If iItem <> 0 Then

        If iItem <> itemIndex Then

            Me.Controls("mybg" & iItem).BackColor = colorWhite

            SwitchColorControl Me.Controls("mybg" & iItem), True

        End If

    End If

    If itemIndex > 0 Then

        Me.Controls("mybg" & itemIndex).BackColor = ColorMove

        iItem = itemIndex

    End If

End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)

    SetControlBackColorUnderline 0

End Sub

Private Sub redBG_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)

    SetControlBackColorUnderline 0

End Sub

Private Sub day11_Click(): SetDayGlobalFromControl 11: End Sub

Private Sub day12_Click(): SetDayGlobalFromControl 12: End Sub

Private Sub day13_Click(): SetDayGlobalFromControl 13: End Sub

Private Sub day14_Click(): SetDayGlobalFromControl 14: End Sub

Private Sub day15_Click(): SetDayGlobalFromControl 15: End Sub

Private Sub day16_Click(): SetDayGlobalFromControl 16: End Sub

Private Sub day17_Click(): SetDayGlobalFromControl 17: End Sub

Private Sub daybg11_Click(): SetDayGlobalFromControl 11: End Sub

Private Sub daybg12_Click(): SetDayGlobalFromControl 12: End Sub

Private Sub daybg13_Click(): SetDayGlobalFromControl 13: End Sub

Private Sub daybg14_Click(): SetDayGlobalFromControl 14: End Sub

Private Sub daybg15_Click(): SetDayGlobalFromControl 15: End Sub

Private Sub daybg16_Click(): SetDayGlobalFromControl 16: End Sub

Private Sub daybg17_Click(): SetDayGlobalFromControl 17: End Sub

Private Sub day21_Click(): SetDayGlobalFromControl 21: End Sub

Private Sub day22_Click(): SetDayGlobalFromControl 22: End Sub

Private Sub day23_Click(): SetDayGlobalFromControl 23: End Sub

Private Sub day24_Click(): SetDayGlobalFromControl 24: End Sub

Private Sub day25_Click(): SetDayGlobalFromControl 25: End Sub

Private Sub day26_Click(): SetDayGlobalFromControl 26: End Sub

Private Sub day27_Click(): SetDayGlobalFromControl 27: End Sub

Private Sub daybg21_Click(): SetDayGlobalFromControl 21: End Sub

Private Sub daybg22_Click(): SetDayGlobalFromControl 22: End Sub

Private Sub daybg23_Click(): SetDayGlobalFromControl 23: End Sub

Private Sub daybg24_Click(): SetDayGlobalFromControl 24: End Sub

Private Sub daybg25_Click(): SetDayGlobalFromControl 25: End Sub

Private Sub daybg26_Click(): SetDayGlobalFromControl 26: End Sub

Private Sub daybg27_Click(): SetDayGlobalFromControl 27: End Sub

Private Sub day31_Click(): SetDayGlobalFromControl 31: End Sub

Private Sub day32_Click(): SetDayGlobalFromControl 32: End Sub

Private Sub day33_Click(): SetDayGlobalFromControl 33: End Sub

Private Sub day34_Click(): SetDayGlobalFromControl 34: End Sub

Private Sub day35_Click(): SetDayGlobalFromControl 35: End Sub

Private Sub day36_Click(): SetDayGlobalFromControl 36: End Sub

Private Sub day37_Click(): SetDayGlobalFromControl 37: End Sub

Private Sub daybg31_Click(): SetDayGlobalFromControl 31: End Sub

Private Sub daybg32_Click(): SetDayGlobalFromControl 32: End Sub

Private Sub daybg33_Click(): SetDayGlobalFromControl 33: End Sub

Private Sub daybg34_Click(): SetDayGlobalFromControl 34: End Sub

Private Sub daybg35_Click(): SetDayGlobalFromControl 35: End Sub

Private Sub daybg36_Click(): SetDayGlobalFromControl 36: End Sub

Private Sub daybg37_Click(): SetDayGlobalFromControl 37: End Sub

Private Sub day41_Click(): SetDayGlobalFromControl 41: End Sub

Private Sub day42_Click(): SetDayGlobalFromControl 42: End Sub

Private Sub day43_Click(): SetDayGlobalFromControl 43: End Sub

Private Sub day44_Click(): SetDayGlobalFromControl 44: End Sub

Private Sub day45_Click(): SetDayGlobalFromControl 45: End Sub

Private Sub day46_Click(): SetDayGlobalFromControl 46: End Sub

Private Sub day47_Click(): SetDayGlobalFromControl 47: End Sub

Private Sub daybg41_Click(): SetDayGlobalFromControl 41: End Sub

Private Sub daybg42_Click(): SetDayGlobalFromControl 42: End Sub

Private Sub daybg43_Click(): SetDayGlobalFromControl 43: End Sub

Private Sub daybg44_Click(): SetDayGlobalFromControl 44: End Sub

Private Sub daybg45_Click(): SetDayGlobalFromControl 45: End Sub

Private Sub daybg46_Click(): SetDayGlobalFromControl 46: End Sub

Private Sub daybg47_Click(): SetDayGlobalFromControl 47: End Sub

Private Sub day51_Click(): SetDayGlobalFromControl 51: End Sub

Private Sub day52_Click(): SetDayGlobalFromControl 52: End Sub

Private Sub day53_Click(): SetDayGlobalFromControl 53: End Sub

Private Sub day54_Click(): SetDayGlobalFromControl 54: End Sub

Private Sub day55_Click(): SetDayGlobalFromControl 55: End Sub

Private Sub day56_Click(): SetDayGlobalFromControl 56: End Sub

Private Sub day57_Click(): SetDayGlobalFromControl 57: End Sub

Private Sub daybg51_Click(): SetDayGlobalFromControl 51: End Sub

Private Sub daybg52_Click(): SetDayGlobalFromControl 52: End Sub

Private Sub daybg53_Click(): SetDayGlobalFromControl 53: End Sub

Private Sub daybg54_Click(): SetDayGlobalFromControl 54: End Sub

Private Sub daybg55_Click(): SetDayGlobalFromControl 55: End Sub

Private Sub daybg56_Click(): SetDayGlobalFromControl 56: End Sub

Private Sub daybg57_Click(): SetDayGlobalFromControl 57: End Sub

Private Sub day61_Click(): SetDayGlobalFromControl 61: End Sub

Private Sub day62_Click(): SetDayGlobalFromControl 62: End Sub

Private Sub day63_Click(): SetDayGlobalFromControl 63: End Sub

Private Sub day64_Click(): SetDayGlobalFromControl 64: End Sub

Private Sub day65_Click(): SetDayGlobalFromControl 65: End Sub

Private Sub day66_Click(): SetDayGlobalFromControl 66: End Sub

Private Sub day67_Click(): SetDayGlobalFromControl 67: End Sub

Private Sub daybg61_Click(): SetDayGlobalFromControl 61: End Sub

Private Sub daybg62_Click(): SetDayGlobalFromControl 62: End Sub

Private Sub daybg63_Click(): SetDayGlobalFromControl 63: End Sub

Private Sub daybg64_Click(): SetDayGlobalFromControl 64: End Sub

Private Sub daybg65_Click(): SetDayGlobalFromControl 65: End Sub

Private Sub daybg66_Click(): SetDayGlobalFromControl 66: End Sub

Private Sub daybg67_Click(): SetDayGlobalFromControl 67: End Sub

Private Sub timeButton_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetDayFromId 68, Button: End Sub

Private Sub todayButton_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetDayFromId 68, Button: End Sub

Private Sub datetimebg_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetDayFromId 68, Button: End Sub

Private Sub day11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 11: End Sub

Private Sub day12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 12: End Sub

Private Sub day13_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 13: End Sub

Private Sub day14_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 14: End Sub

Private Sub day15_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 15: End Sub

Private Sub day16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 16: End Sub

Private Sub day17_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 17: End Sub

Private Sub daybg11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 11: End Sub

Private Sub daybg12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 12: End Sub

Private Sub daybg13_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 13: End Sub

Private Sub daybg14_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 14: End Sub

Private Sub daybg15_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 15: End Sub

Private Sub daybg16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 16: End Sub

Private Sub daybg17_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 17: End Sub

Private Sub day21_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 21: End Sub

Private Sub day22_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 22: End Sub

Private Sub day23_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 23: End Sub

Private Sub day24_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 24: End Sub

Private Sub day25_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 25: End Sub

Private Sub day26_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 26: End Sub

Private Sub day27_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 27: End Sub

Private Sub daybg21_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 21: End Sub

Private Sub daybg22_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 22: End Sub

Private Sub daybg23_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 23: End Sub

Private Sub daybg24_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 24: End Sub

Private Sub daybg25_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 25: End Sub

Private Sub daybg26_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 26: End Sub

Private Sub daybg27_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 27: End Sub

Private Sub day31_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 31: End Sub

Private Sub day32_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 32: End Sub

Private Sub day33_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 33: End Sub

Private Sub day34_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 34: End Sub

Private Sub day35_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 35: End Sub

Private Sub day36_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 36: End Sub

Private Sub day37_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 37: End Sub

Private Sub daybg31_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 31: End Sub

Private Sub daybg32_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 32: End Sub

Private Sub daybg33_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 33: End Sub

Private Sub daybg34_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 34: End Sub

Private Sub daybg35_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 35: End Sub

Private Sub daybg36_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 36: End Sub

Private Sub daybg37_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 37: End Sub

Private Sub day41_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 41: End Sub

Private Sub day42_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 42: End Sub

Private Sub day43_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 43: End Sub

Private Sub day44_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 44: End Sub

Private Sub day45_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 45: End Sub

Private Sub day46_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 46: End Sub

Private Sub day47_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 47: End Sub

Private Sub daybg41_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 41: End Sub

Private Sub daybg42_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 42: End Sub

Private Sub daybg43_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 43: End Sub

Private Sub daybg44_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 44: End Sub

Private Sub daybg45_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 45: End Sub

Private Sub daybg46_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 46: End Sub

Private Sub daybg47_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 47: End Sub

Private Sub day51_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 51: End Sub

Private Sub day52_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 52: End Sub

Private Sub day53_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 53: End Sub

Private Sub day54_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 54: End Sub

Private Sub day55_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 55: End Sub

Private Sub day56_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 56: End Sub

Private Sub day57_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 57: End Sub

Private Sub daybg51_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 51: End Sub

Private Sub daybg52_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 52: End Sub

Private Sub daybg53_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 53: End Sub

Private Sub daybg54_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 54: End Sub

Private Sub daybg55_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 55: End Sub

Private Sub daybg56_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 56: End Sub

Private Sub daybg57_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 57: End Sub

Private Sub day61_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 61: End Sub

Private Sub day62_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 62: End Sub

Private Sub day63_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 63: End Sub

Private Sub day64_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 64: End Sub

Private Sub day65_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 65: End Sub

Private Sub day66_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 66: End Sub

Private Sub day67_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 67: End Sub

Private Sub daybg61_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 61: End Sub

Private Sub daybg62_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 62: End Sub

Private Sub daybg63_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 63: End Sub

Private Sub daybg64_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 64: End Sub

Private Sub daybg65_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 65: End Sub

Private Sub daybg66_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 66: End Sub

Private Sub daybg67_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 67: End Sub

Private Sub datetimebg_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 68: End Sub

Private Sub monthTitle_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 69: End Sub

Private Sub yearTitle_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColorUnderline 70: End Sub

Private Sub my1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColor 1: End Sub

Private Sub my2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColor 2: End Sub

Private Sub my3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColor 3: End Sub

Private Sub my4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColor 4: End Sub

Private Sub my5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColor 5: End Sub

Private Sub my6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColor 6: End Sub

Private Sub my7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColor 7: End Sub

Private Sub my8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColor 8: End Sub

Private Sub my9_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColor 9: End Sub

Private Sub my10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColor 10: End Sub

Private Sub my11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColor 11: End Sub

Private Sub my12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColor 12: End Sub

Private Sub mybg1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColor 1: End Sub

Private Sub mybg2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColor 2: End Sub

Private Sub mybg3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColor 3: End Sub

Private Sub mybg4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColor 4: End Sub

Private Sub mybg5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColor 5: End Sub

Private Sub mybg6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColor 6: End Sub

Private Sub mybg7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColor 7: End Sub

Private Sub mybg8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColor 8: End Sub

Private Sub mybg9_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColor 9: End Sub

Private Sub mybg10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColor 10: End Sub

Private Sub mybg11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColor 11: End Sub

Private Sub mybg12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColor 12: End Sub

Private Sub myFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single): SetControlBackColor 0: End Sub

Private Sub my1_Click(): SetMountYearFormIdControl 1: End Sub

Private Sub my2_Click(): SetMountYearFormIdControl 2: End Sub

Private Sub my3_Click(): SetMountYearFormIdControl 3: End Sub

Private Sub my4_Click(): SetMountYearFormIdControl 4: End Sub

Private Sub my5_Click(): SetMountYearFormIdControl 5: End Sub

Private Sub my6_Click(): SetMountYearFormIdControl 6: End Sub

Private Sub my7_Click(): SetMountYearFormIdControl 7: End Sub

Private Sub my8_Click(): SetMountYearFormIdControl 8: End Sub

Private Sub my9_Click(): SetMountYearFormIdControl 9: End Sub

Private Sub my10_Click(): SetMountYearFormIdControl 10: End Sub

Private Sub my11_Click(): SetMountYearFormIdControl 11: End Sub

Private Sub my12_Click(): SetMountYearFormIdControl 12: End Sub

Private Sub mybg1_Click(): SetMountYearFormIdControl 1: End Sub

Private Sub mybg2_Click(): SetMountYearFormIdControl 2: End Sub

Private Sub mybg3_Click(): SetMountYearFormIdControl 3: End Sub

Private Sub mybg4_Click(): SetMountYearFormIdControl 4: End Sub

Private Sub mybg5_Click(): SetMountYearFormIdControl 5: End Sub

Private Sub mybg6_Click(): SetMountYearFormIdControl 6: End Sub

Private Sub mybg7_Click(): SetMountYearFormIdControl 7: End Sub

Private Sub mybg8_Click(): SetMountYearFormIdControl 8: End Sub

Private Sub mybg9_Click(): SetMountYearFormIdControl 9: End Sub

Private Sub mybg10_Click(): SetMountYearFormIdControl 10: End Sub

Private Sub mybg11_Click(): SetMountYearFormIdControl 11: End Sub

Private Sub mybg12_Click(): SetMountYearFormIdControl 12: End Sub

'-------------------------------------------------------------------------

Private Sub SwitchColorControl(objControl As control, Optional picker As Boolean = False)

    If objControl.BackColor = ColorSelectedDate Then

        objControl.BackColor = colorWhite

    End If

    If picker Then

        If iSelectGlobal = objControl.Tag Then

            objControl.BackColor = ColorSelectedDate

        End If

    Else

        If DateGlobal = objControl.Tag Then

            objControl.BackColor = ColorSelectedDate

        End If

    End If

End Sub

Private Sub SetCapitonControlSunday()

    Dim i           As Byte

    For i = 1 To 7

        Me.Controls("dayofweek" & i).Caption = VBA.WeekdayName(i, True, vbMonday)

    Next i

End Sub



Private Sub SetDayFromId(id As Integer, Button As Integer)

    If Button > 1 Then

        SetDayGlobalFromControl 68, True

    Else

        SetDayGlobalFromControl 68

    End If

End Sub

Private Sub SetDayGlobalFromControl(id As Integer, Optional inculdeTime As Boolean = False)

    If id <= 67 Then

        sDateGlobal = Me.Controls("day" & id).Tag

    ElseIf id = 68 Then

        sDateGlobal = VBA.Date

        If inculdeTime Then sDateGlobal = sDateGlobal & " " & VBA.Time

    End If

    lbDateChose.Caption = sDateGlobal

    Me.Hide

End Sub

Private Sub SetMountYearFormIdControl(id As Integer)

    If lbTypeCalendar.Caption = "1" Then

        lbDateChose.Caption = Me.Controls("my" & id).Tag

        Me.Hide

        Exit Sub

    End If

    If pickerMode = 1 Then

        iMounthGlobal = Me.Controls("my" & id).Tag

    ElseIf pickerMode = 2 Then

        iYearGlobal = Me.Controls("my" & id).Tag



        If minDate <> 0 And iYearGlobal = VBA.Year(minDate) Then

            If iMounthGlobal < VBA.Month(minDate) Then iMounthGlobal = VBA.Month(minDate)

        End If



        If maxDate <> 0 And iYearGlobal = VBA.Year(maxDate) Then

            If iMounthGlobal > VBA.Month(maxDate) Then iMounthGlobal = VBA.Month(maxDate)

        End If

    End If

    SetDayOneAndDayTwo

    SwitchVisibleButton 0

End Sub



Private Sub SetShowPicerMode1()

    If pickerMode = 1 Then

        SwitchVisibleButton 0

    Else

        SwitchVisibleButton 1

        FrameTop30

    End If

End Sub



Private Sub SetShowPicerMode2()

    If pickerMode = 2 Then

        SwitchVisibleButton 0

    Else

        SwitchVisibleButton 2

        FrameTop30

    End If

End Sub



Private Sub FrameTop30()

    myFrame.Top = 30

End Sub



Private Sub FrameTopHeight()

    myFrame.Top = Me.Height

End Sub



Private Sub SwitchVisibleButton(Mode As Integer)

    If Mode = 0 Then

        pickerMode = 0

        FrameTopHeight

        Call subVisibleButtonMount(True)

    ElseIf Mode = 1 Then

        Call SetMounts

        Call subVisibleButtonMount(False)

        pickerMode = 1

        Call FrameTop30

    ElseIf Mode = 2 Then

        iShiftGlobal = -6

        Call SetYears

        Call subVisibleButtonMount(True)

        pickerMode = 2

        Call FrameTop30

    End If

End Sub



Private Sub blockBtn(Mode As Integer)

    If Mode = 0 Then

        prevMonthButton.Enabled = day11.Enabled

        nextMonthButton.Enabled = day67.Enabled

    ElseIf Mode = 1 Then

        prevMonthButton.Enabled = myBG1.Enabled

        nextMonthButton.Enabled = myBG12.Enabled

    ElseIf Mode = 2 Then



    End If

End Sub



Private Sub SetDayOneAndDayTwo()

    Dim dtDate1     As Date

    Dim dtDate2     As Date

    Dim iDay        As Integer

    Dim objCnt1     As control

    Dim objCnt2     As control

    Dim i           As Byte

    Dim j           As Byte



    dtDate1 = VBA.DateSerial(iYearGlobal, iMounthGlobal, 1)

    iDay = VBA.Weekday(dtDate1, vbMonday)

    dtDate2 = VBA.DateAdd("d", -iDay + 1, dtDate1)

    Me.monthTitle.Caption = VBA.MonthName(VBA.Month(dtDate1), True)

    Me.yearTitle.Caption = VBA.Year(dtDate1)

    For i = 1 To 6

        For j = 1 To 7

            Set objCnt1 = Me.Controls("day" & i & j)

            Set objCnt2 = Me.Controls("dayBG" & i & j)

            objCnt1.Caption = VBA.day(dtDate2)

            objCnt1.Tag = dtDate2

            objCnt2.Tag = dtDate2





            objCnt1.Enabled = Not ((dtDate2 < minDate And minDate <> 0) Or (dtDate2 > maxDate And maxDate <> 0))

            objCnt2.Enabled = objCnt1.Enabled



            If VBA.Month(dtDate2) <> iMounthGlobal Then

                objCnt1.ForeColor = ForeColorDayNo

            Else

                objCnt1.ForeColor = ForeColorDayYes

            End If

            SwitchColorControl objCnt2

            dtDate2 = VBA.DateAdd("d", 1, dtDate2)

        Next j

    Next i

    Call blockBtn(0)

End Sub



Private Sub SetMounts()

    Dim objCnt1     As control

    Dim objCnt2     As control

    Dim i           As Byte

    Dim minMonth    As Byte

    Dim maxMonth    As Byte

    Dim minYear     As Integer

    Dim maxYear     As Integer



    minMonth = VBA.Month(minDate)

    maxMonth = VBA.Month(maxDate)

    minYear = VBA.Year(minDate)

    maxYear = VBA.Year(maxDate)

    For i = 1 To 12

        Set objCnt1 = Me.Controls("my" & i)

        Set objCnt2 = Me.Controls("mybg" & i)

        objCnt1.Caption = VBA.MonthName(i, True)

        objCnt1.Tag = i

        objCnt2.Tag = i



        objCnt1.Enabled = Not ((i < minMonth And minDate <> 0 And iYearGlobal = minYear) Or (i > maxMonth And maxDate <> 0 And iYearGlobal = maxYear))

        objCnt2.Enabled = objCnt1.Enabled



        objCnt2.BackColor = colorWhite

        SwitchColorControl objCnt2, True

    Next i

End Sub



Private Sub SetYears()

    Dim objCnt1     As control

    Dim objCnt2     As control

    Dim iYear       As Integer

    Dim i           As Byte



    iYear = iYearGlobal + iShiftGlobal

    For i = 1 To 12

        Set objCnt1 = Me.Controls("my" & i)

        Set objCnt2 = Me.Controls("mybg" & i)

        objCnt1.Caption = iYear

        objCnt1.Tag = iYear

        objCnt2.Tag = iYear



        objCnt1.Enabled = Not ((iYear < VBA.Year(minDate) And minDate <> 0) Or (iYear > VBA.Year(maxDate) And maxDate <> 0))

        objCnt2.Enabled = objCnt1.Enabled

        objCnt2.BackColor = colorWhite

        SwitchColorControl objCnt2, True

        iYear = iYear + 1

    Next i

    Call blockBtn(1)

End Sub



Private Sub subVisibleButtonMount(ByVal Value As Boolean)

    prevMonthButton.Visible = Value

    nextMonthButton.Visible = Value

End Sub
