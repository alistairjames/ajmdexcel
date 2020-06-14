Attribute VB_Name = "PeakSeeker"
'PROGRAMME FOR FINDING PEAK CENTRE GIVEN START AND END
'AND CALCULATING THE WIDTH AT HALF HEIGHT AND
'DISPLAYING THE DATA IN A TABLE AND ON A GRAPH

'A SERIES OF VARIABLES NEEDED BY THE PROGRAMME
Private TimePoint As Integer
Private Height As Single
Private PeakTime As Single                ' value in minutes
Private PeakTimePoint As Single           'the row containing the peak time data
Private PeakStartTime As Single           'obtained from the spreadsheet
Private PeakStartTimePoint As Integer     'calculated from the PeakStartTime
Private PeakStartRI As Single
Private PeakEndTime As Single
Private PeakEndTimePoint As Integer
Private PeakEndRI As Single
Private Van(4) As Single
Private Rear(4) As Single
Private HalfHeight As Single
Private HalfStartTimePoint As Integer
Private HalfEndTimePoint As Integer
Private HalfHeightWidth As Single
Private HeightByWidth As Single


Sub ImportData() 'triggered from macro button on sheet

'Clears the old data and the results
Range(Cells(11, 1), Cells(12, 56)).ClearContents
Range(Cells(13, 3), Cells(5000, 7)).ClearContents
Range(Cells(4, 6), Cells(9, 6)).ClearContents
Cells(4, 2).ClearContents
Cells(5, 2).ClearContents
DoTheImport

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DoTheImport
' This prompts the user for a FileName as separator character
' and then calls ImportTextFile.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DoTheImport()
    Dim FileName As Variant
    Dim Sep As String
    FileName = Application.GetOpenFilename
    If FileName = False Then
        ''''''''''''''''''''''''''
        ' user cancelled, get out
        ''''''''''''''''''''''''''
        Exit Sub
    End If
    Sep = vbTab
    Debug.Print "FileName: " & FileName, "Separator: " & Sep
    ImportTextFile FName:=CStr(FileName), Sep:=CStr(Sep)
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ImportTextFile
' This imports a text file into Excel.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ImportTextFile(FName As String, Sep As String)

Dim RowNdx As Long
Dim ColNdx As Integer
Dim TempVal As Variant
Dim WholeLine As String
Dim Pos As Integer
Dim NextPos As Integer
Dim SaveColNdx As Integer

Application.ScreenUpdating = False
'On Error GoTo EndMacro:

Cells(3, 2).Value = FName    'Added by me to display the FileName
SaveColNdx = 1               'Sets the input to column 1 (example used ActiveCell.Column)
RowNdx = 11                  'Sets the input row to 11 (example used ActiveCell.Row)

Open FName For Input Access Read As #1

While Not EOF(1)
    Line Input #1, WholeLine
    If Right(WholeLine, 1) <> Sep Then   ' in vb <> means not equals
        WholeLine = WholeLine & Sep
    End If
    ColNdx = SaveColNdx
    Pos = 1
    NextPos = InStr(Pos, WholeLine, Sep)
    While NextPos >= 1
        TempVal = Mid(WholeLine, Pos, NextPos - Pos)
        Cells(RowNdx, ColNdx).Value = TempVal          'where the data is pasted in
        Pos = NextPos + 1
        ColNdx = ColNdx + 1
        NextPos = InStr(Pos, WholeLine, Sep)
    Wend
    RowNdx = RowNdx + 1
Wend

EndMacro:
On Error GoTo 0
Application.ScreenUpdating = True
Close #1

End Sub

'Finds the peak based on user input of start and end points into the spreadsheet
'Triggered from button on worksheet
Sub PeakSeeker()

'Stops the macro running if there is no user input
If IsEmpty(Cells(4, 2)) Or IsEmpty(Cells(5, 2)) Then
  MsgBox ("Peak start or end times are missing in Cells B4 and B5. ")
  Exit Sub
End If

'Need to clear any old data imported or calculated and get user input
Range(Cells(13, 3), Cells(5000, 7)).ClearContents
PeakStartTime = Cells(4, 2).Value
PeakEndTime = Cells(5, 2).Value

'Calculate the TimePoints for peak start and end.
' TimePoint set at 1 min to start with
TimePoint = 60
Do
TimePoint = TimePoint + 1
Loop Until Cells(TimePoint, 1).Value > PeakStartTime

PeakStartTimePoint = TimePoint

Do
TimePoint = TimePoint + 1
Loop Until Cells(TimePoint, 1).Value > PeakEndTime

PeakEndTimePoint = TimePoint

'Locate the peak centre
TimePoint = PeakStartTimePoint
FindPeakCentre
  Cells(4, 6).Value = PeakTimePoint
  Cells(5, 6).Value = PeakTime
  
DrawBaseline
Height = Cells(PeakTimePoint, 2).Value - Cells(PeakTimePoint, 3).Value
Cells(6, 6).Value = Height
HalfHeight = Height / 2

'Draw peak marker
Cells(PeakTimePoint, 5).Value = Cells(PeakTimePoint, 1).Value
Cells(PeakTimePoint, 6).Value = Cells(PeakTimePoint, 2).Value
Cells(PeakTimePoint + 1, 5).Value = Cells(PeakTimePoint, 1).Value
Cells(PeakTimePoint + 1, 6).Value = Cells(PeakTimePoint, 3).Value

''''''Calculate Width at Half-Height
'Find first Timepoint inside peak that is greater than HalfHeight
TimePoint = PeakStartTimePoint
Do
TimePoint = TimePoint + 1
Loop Until Cells(TimePoint, 2) - Cells(TimePoint, 3) > HalfHeight
HalfStartTimePoint = TimePoint

'Find last TimePoint inside peak that is greater than HalfHeight
TimePoint = PeakEndTimePoint
Do
TimePoint = TimePoint - 1
Loop Until Cells(TimePoint, 2) - Cells(TimePoint, 3) > HalfHeight
HalfEndTimePoint = TimePoint

'The main part of the line is measured less the little bits ot be added
HalfHeightWidth = HalfEndTimePoint - HalfStartTimePoint 'implicit conversion to Single. Units are TimePoints

'INTERPOLATION BETWEEN TIMEPOINTS TO GET THE EXACT PEAK WIDTH AT HALF HEIGHT
'Calculate how much time to add before and after the TimePoints identified above
' Calculation done in TimePoints and then converted to minutes from the difference between TimePoints
'All height calculations are done from the drawn baseline, so the figure is sheared to give a right-angled triangle
' h - excess height of HalfStartTimePoint over Previous TimePoint
' d - excess of HalfHeight over height at the Previous TimePoint
' p - the horizontal distance (in TimePoints) from previous TimePoint to the place where the true HalfHeight time is.
' alpha - the angle made by the ascending curve and the horizontal
' tan alpha is h/1 ......  tan alpha is also d/p   ....   so h = d/p   .......  and p = d/h

'For the start of the half-height line
Dim h As Single
Dim d As Single
Dim p As Single

h = (Cells(HalfStartTimePoint, 2).Value - Cells(HalfStartTimePoint, 3).Value) - _
                           (Cells(HalfStartTimePoint - 1, 2).Value - Cells(HalfStartTimePoint - 1, 3).Value)

d = HalfHeight - (Cells(HalfStartTimePoint - 1, 2).Value - Cells(HalfStartTimePoint - 1, 3).Value)
p = d / h

'increment the half height width by this little bit
HalfHeightWidth = HalfHeightWidth + (1 - p) ' p is the distance before the Width starts so add 1-p

'Repeat for the end of the half-height line
'The calculation is the same as before only flipped left to right
h = (Cells(HalfEndTimePoint, 2).Value - Cells(HalfEndTimePoint, 3).Value) - _
                           (Cells(HalfEndTimePoint + 1, 2).Value - Cells(HalfEndTimePoint + 1, 3).Value)

d = HalfHeight - (Cells(HalfEndTimePoint + 1, 2).Value - Cells(HalfEndTimePoint + 1, 3).Value)
p = d / h

'increment the half height width for this little bit
HalfHeightWidth = HalfHeightWidth + (1 - p) ' p is the distance after the Width ends so add 1-p

DrawHalfHeightLine

'Report width and ratio to height.
Cells(7, 6).Value = HalfHeightWidth  'in TimePoints

'Peak width conversion of units from TimePoints to minutes and ratio to peak height
'Time column units are assumed to be minutes
'Two arbitrarily selected successive cells are used to get the time difference between TimePoints
HeightByWidth = Height / (HalfHeightWidth * (Cells(51, 1).Value - Cells(50, 1).Value))
Cells(9, 6).Value = HeightByWidth


End Sub


Sub FindPeakCentre()

' move along to start of peak - gets past areas of baseline wobble
' gets the gradient from 7 TimePoints before and behind the current TimePoint
Do
TimePoint = TimePoint + 1
Loop Until (Cells(TimePoint + 7, 2).Value - Cells(TimePoint - 7, 2).Value) / 15 > 0.025

PopulateTrain
'TimePoint is the centre of the train
'Rear section is TimePoint -7,-6,-5,-4
'Van section is Timepoint +4, +5, +6, +7
'This gives a bit of averaging to the gradient estimate

Do While Van(0) + Van(1) + Van(2) + Van(3) > Rear(0) + Rear(1) + Rear(2) + Rear(3)
   TimePoint = TimePoint + 1
   PopulateTrain
  Loop

Height = Cells(TimePoint, 2).Value
PeakTime = Cells(TimePoint, 1).Value
PeakTimePoint = TimePoint

  
End Sub



Sub DrawBaseline()

PeakStartRI = Cells(PeakStartTimePoint, 2).Value
PeakEndRI = Cells(PeakEndTimePoint, 2).Value
For TimePoint = PeakStartTimePoint To PeakEndTimePoint
     Cells(TimePoint, 3).Value = PeakStartRI + (PeakEndRI - PeakStartRI) * (TimePoint - PeakStartTimePoint) _
                                                      / (PeakEndTimePoint - PeakStartTimePoint)

Next TimePoint

End Sub


'Draws a line parralel to the baseline at the HalfHeight just inside the existing trace

Sub DrawHalfHeightLine()

For TimePoint = HalfStartTimePoint To HalfEndTimePoint
     Cells(TimePoint, 4).Value = Cells(TimePoint, 3).Value + HalfHeight

Next TimePoint



End Sub


Sub PopulateTrain()
Rear(0) = Cells(TimePoint - 7, 2).Value 'Get numbers into Rear array
Rear(1) = Cells(TimePoint - 6, 2).Value
Rear(2) = Cells(TimePoint - 5, 2).Value
Rear(3) = Cells(TimePoint - 4, 2).Value

Van(0) = Cells(TimePoint + 4, 2).Value 'Get numbers into Van array
Van(1) = Cells(TimePoint + 5, 2).Value
Van(2) = Cells(TimePoint + 6, 2).Value
Van(3) = Cells(TimePoint + 7, 2).Value
End Sub


