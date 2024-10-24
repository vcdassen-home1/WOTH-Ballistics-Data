Attribute VB_Name = "Test_TrajectoryTracking"
'@TestModule
'@Folder("Tests")


Option Explicit
Option Private Module

Private Assert As Rubberduck.AssertClass
Private Fakes As Rubberduck.FakesProvider

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
    Set Fakes = New Rubberduck.FakesProvider
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("CTrajectoryTracker")
Private Sub Test_TrajectoryTracker_Construction()
    On Error GoTo TestFail
    Const nPoints As Long = 10
    Const yMin As Double = 0#
    Const yMax As Double = 10#
    
    'Arrange:
    Dim tracker As CTrajectoryTracker
    
    'Act:
    Set tracker = TrajectoryTracker(nPoints, yMin, yMax)
    
    'Assert:
    Assert.IsNotNothing tracker, "tracker constructed"
    Assert.AreEqual nPoints, tracker.Size, "tracker size = 10"
    Assert.AreEqual yMin, tracker.yMin
    Assert.AreEqual yMax, tracker.yMax

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("CTrajectoryTracker")
Private Sub Test_TrajectoryTracker_TrackTrajectory()
    On Error GoTo TestFail
    Const nPoints As Long = 10
    Const yMin As Double = 0#
    Const yMax As Double = 2#
    
    Dim trajectoryPoints() As Double
    ReDim trajectoryPoints(1 To nPoints, 1 To 2)
    
    
    'Arrange:
    trajectoryPoints(1, 1) = 0#
    trajectoryPoints(1, 2) = -1#  ' below target
    
    trajectoryPoints(2, 1) = 0.5
    trajectoryPoints(2, 2) = 0#   ' inside target
    
    trajectoryPoints(3, 1) = 1#
    trajectoryPoints(3, 2) = 1.5  ' inside target
    
    trajectoryPoints(4, 1) = 2#
    trajectoryPoints(4, 2) = 2.5   'above target
    
    trajectoryPoints(5, 1) = 2.5
    trajectoryPoints(5, 2) = 3#    'above target
    
    trajectoryPoints(6, 1) = 3.5
    trajectoryPoints(6, 2) = 4.5   'above target
    
    trajectoryPoints(7, 1) = 4#
    trajectoryPoints(7, 2) = 2.5   'above target
    
    trajectoryPoints(8, 1) = 4.5
    trajectoryPoints(8, 2) = 2.1    'above target
    
    trajectoryPoints(9, 1) = 5#
    trajectoryPoints(9, 2) = 1.5    'inside target
    
    trajectoryPoints(10, 1) = 5.5
    trajectoryPoints(10, 2) = -0.5  'below target
    
    
    'Act:
    Dim tracker As CTrajectoryTracker
    Set tracker = TrajectoryTracker(nPoints, yMin, yMax)
    Dim i As Long
    For i = LBound(trajectoryPoints, 1) To UBound(trajectoryPoints, 1)
        tracker.Track trajectoryPoints(i, 1), trajectoryPoints(i, 2)
    Next i
    
    'Assert:
    Assert.AreEqual nPoints, UBound(tracker.TrajectoryData, 1)
    Assert.AreEqual nPoints, UBound(tracker.TargetCrossings)
    
    Dim crossings() As ETargetDiameterCrossing: crossings = tracker.TargetCrossings
    Assert.AreEqual ETargetDiameterCrossing.BelowTarget, crossings(1)
    Assert.AreEqual ETargetDiameterCrossing.InsideTarget, crossings(2)
    Assert.AreEqual ETargetDiameterCrossing.InsideTarget, crossings(3)
    Assert.AreEqual ETargetDiameterCrossing.AboveTarget, crossings(4)
    Assert.AreEqual ETargetDiameterCrossing.AboveTarget, crossings(5)
    Assert.AreEqual ETargetDiameterCrossing.AboveTarget, crossings(6)
    Assert.AreEqual ETargetDiameterCrossing.AboveTarget, crossings(7)
    Assert.AreEqual ETargetDiameterCrossing.AboveTarget, crossings(8)
    Assert.AreEqual ETargetDiameterCrossing.InsideTarget, crossings(9)
    Assert.AreEqual ETargetDiameterCrossing.BelowTarget, crossings(10)

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("CTrajectoryTracker")
Private Sub Test_TrajectoryTracker_FindCrossingPoints()
    On Error GoTo TestFail
    Const nPoints As Long = 10
    Const yMin As Double = 0#
    Const yMax As Double = 2#
    Const nCrossingPoints As Long = 4
    
    Dim trajectoryPoints() As Double
    ReDim trajectoryPoints(1 To nPoints, 1 To 2)
    
    
    'Arrange:
    trajectoryPoints(1, 1) = 0#
    trajectoryPoints(1, 2) = -1#  ' below target
    
    trajectoryPoints(2, 1) = 0.5
    trajectoryPoints(2, 2) = 0#   ' inside target first crossing
    
    trajectoryPoints(3, 1) = 1#
    trajectoryPoints(3, 2) = 1.5  ' inside target
    
    trajectoryPoints(4, 1) = 2#
    trajectoryPoints(4, 2) = 2.5   'above target second crossing
    
    trajectoryPoints(5, 1) = 2.5
    trajectoryPoints(5, 2) = 3#    'above target
    
    trajectoryPoints(6, 1) = 3.5
    trajectoryPoints(6, 2) = 4.5   'above target
    
    trajectoryPoints(7, 1) = 4#
    trajectoryPoints(7, 2) = 2.5   'above target
    
    trajectoryPoints(8, 1) = 4.5
    trajectoryPoints(8, 2) = 2.1    'above target
    
    trajectoryPoints(9, 1) = 5#
    trajectoryPoints(9, 2) = 1.5    'inside target third crossing
    
    trajectoryPoints(10, 1) = 5.5
    trajectoryPoints(10, 2) = -0.5  'below target fourth crossing
    
    
    'Act:
    Dim tracker As CTrajectoryTracker
    Set tracker = TrajectoryTracker(nPoints, yMin, yMax)
    Dim i As Long
    For i = LBound(trajectoryPoints, 1) To UBound(trajectoryPoints, 1)
        tracker.Track trajectoryPoints(i, 1), trajectoryPoints(i, 2)
    Next i
    
    Dim trackingList As New ArrayList
    tracker.FindCrossings trackingList
    
    
    
    'Assert:
    Assert.AreEqual nCrossingPoints, trackingList.Count
   

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("CTrajectoryTracker")
Private Sub Test_TrajectoryTracker_PBR_ZeroCrossings()
    On Error GoTo TestFail
    Const nPoints As Long = 10
    Const yMin As Double = -2#
    Const yMax As Double = 2#
    Const xStep As Double = 0.5
    Const nCrossingPoints As Long = 0
    Const nPBRIntervals As Long = 0
    
    Dim trajectoryPoints() As Double
    ReDim trajectoryPoints(1 To nPoints, 1 To 2)
    
    
    'Arrange:
    Dim i As Long
    Dim x As Double
    Dim y As Double
    
    For i = LBound(trajectoryPoints, 1) To UBound(trajectoryPoints, 1)
        x = xStep * i
        trajectoryPoints(i, 1) = x
        y = yMax + 2# * x
        trajectoryPoints(i, 2) = y
    Next i
    
    
    'Act:
    Dim tracker As CTrajectoryTracker
    Set tracker = TrajectoryTracker(nPoints, yMin, yMax)
    
    For i = LBound(trajectoryPoints, 1) To UBound(trajectoryPoints, 1)
        tracker.Track trajectoryPoints(i, 1), trajectoryPoints(i, 2)
    Next i
    
    Dim trackingList As New ArrayList
    tracker.FindCrossings trackingList
    
    Dim pbrDataList As New ArrayList
    tracker.BuildPointBlankRangeList trackingList, pbrDataList
    
    'Assert:
    Assert.AreEqual nCrossingPoints, trackingList.Count
    Assert.AreEqual nPBRIntervals, pbrDataList.Count
   

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("CTrajectoryTracker")
Private Sub Test_TrajectoryTracker_PBR_OneCrossing()
    On Error GoTo TestFail
    Const nPoints As Long = 25
    Const yMin As Double = -2#
    Const yMax As Double = 2#
    Const xStep As Double = 0.1
    Const nCrossingPoints As Long = 1
    Const nPBRIntervals As Long = 1
    
    Dim trajectoryPoints() As Double
    ReDim trajectoryPoints(1 To nPoints, 1 To 2)
    
    
    'Arrange:
    Dim i As Long
    Dim x As Double
    Dim y As Double
    
    For i = LBound(trajectoryPoints, 1) To UBound(trajectoryPoints, 1)
        x = xStep * i
        trajectoryPoints(i, 1) = x
        y = yMin + 0.2 + 4# * x - 2# * x * x
        trajectoryPoints(i, 2) = y
    Next i
    
    
    'Act:
    Dim tracker As CTrajectoryTracker
    Set tracker = TrajectoryTracker(nPoints, yMin, yMax)
    
    For i = LBound(trajectoryPoints, 1) To UBound(trajectoryPoints, 1)
        tracker.Track trajectoryPoints(i, 1), trajectoryPoints(i, 2)
    Next i
    
    Dim trackingList As New ArrayList
    tracker.FindCrossings trackingList
    
    Dim pbrDataList As New ArrayList
    tracker.BuildPointBlankRangeList trackingList, pbrDataList
    
    'Assert:
    Assert.AreEqual nCrossingPoints, trackingList.Count
    Assert.AreEqual nPBRIntervals, pbrDataList.Count
   

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("CTrajectoryTracker")
Private Sub Test_TrajectoryTracker_PBR_TwoCrossingsCase1()
    On Error GoTo TestFail
    Const nPoints As Long = 25
    Const yMin As Double = -2#
    Const yMax As Double = 2#
    Const xStep As Double = 0.1
    Const nCrossingPoints As Long = 2
    Const nPBRIntervals As Long = 1
    
    Dim trajectoryPoints() As Double
    ReDim trajectoryPoints(1 To nPoints, 1 To 2)
    
    
    'Arrange:
    Dim i As Long
    Dim x As Double
    Dim y As Double
    
    For i = LBound(trajectoryPoints, 1) To UBound(trajectoryPoints, 1)
        x = xStep * i
        trajectoryPoints(i, 1) = x
        y = yMin - 1# + 4# * x - 2# * x * x
        trajectoryPoints(i, 2) = y
    Next i
    
    
    'Act:
    Dim tracker As CTrajectoryTracker
    Set tracker = TrajectoryTracker(nPoints, yMin, yMax)
    
    For i = LBound(trajectoryPoints, 1) To UBound(trajectoryPoints, 1)
        tracker.Track trajectoryPoints(i, 1), trajectoryPoints(i, 2)
    Next i
    
    Dim trackingList As New ArrayList
    tracker.FindCrossings trackingList
    
    Dim pbrDataList As New ArrayList
    tracker.BuildPointBlankRangeList trackingList, pbrDataList
    
    'Assert:
    Assert.AreEqual nCrossingPoints, trackingList.Count
    Assert.AreEqual nPBRIntervals, pbrDataList.Count
   

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("CTrajectoryTracker")
Private Sub Test_TrajectoryTracker_PBR_TwoCrossingsCase2()
    On Error GoTo TestFail
    Const nPoints As Long = 25
    Const yMin As Double = -2#
    Const yMax As Double = 2#
    Const xStep As Double = 0.1
    Const nCrossingPoints As Long = 2
    Const nPBRIntervals As Long = 1
    
    Dim trajectoryPoints() As Double
    ReDim trajectoryPoints(1 To nPoints, 1 To 2)
    
    
    'Arrange:
    Dim i As Long
    Dim x As Double
    Dim y As Double
    
    For i = LBound(trajectoryPoints, 1) To UBound(trajectoryPoints, 1)
        x = xStep * i
        trajectoryPoints(i, 1) = x
        y = yMax + 1# + 0.5 * x - 2 * x * x
        trajectoryPoints(i, 2) = y
    Next i
    
    
    'Act:
    Dim tracker As CTrajectoryTracker
    Set tracker = TrajectoryTracker(nPoints, yMin, yMax)
    
    For i = LBound(trajectoryPoints, 1) To UBound(trajectoryPoints, 1)
        tracker.Track trajectoryPoints(i, 1), trajectoryPoints(i, 2)
    Next i
    
    Dim trackingList As New ArrayList
    tracker.FindCrossings trackingList
    
    Dim pbrDataList As New ArrayList
    tracker.BuildPointBlankRangeList trackingList, pbrDataList
    
    'Assert:
    Assert.AreEqual nCrossingPoints, trackingList.Count
    Assert.AreEqual nPBRIntervals, pbrDataList.Count
   

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub




'@TestMethod("CTrajectoryTracker")
Private Sub Test_TrajectoryTracker_FourCrossings()
    On Error GoTo TestFail
    Const nPoints As Long = 11
    Const yMin As Double = 0#
    Const yMax As Double = 2#
    Const nCrossingPoints As Long = 4
    Const nPBRIntervals As Long = 2
    
    Dim trajectoryPoints() As Double
    ReDim trajectoryPoints(1 To nPoints, 1 To 2)
    
    
    'Arrange:
    trajectoryPoints(1, 1) = 0#
    trajectoryPoints(1, 2) = -1#  ' below target
    
    trajectoryPoints(2, 1) = 0.5
    trajectoryPoints(2, 2) = 0#   ' inside target first crossing
    
    trajectoryPoints(3, 1) = 1#
    trajectoryPoints(3, 2) = 1.5  ' inside target
    
    trajectoryPoints(4, 1) = 2#
    trajectoryPoints(4, 2) = 2.5   'above target second crossing
    
    trajectoryPoints(5, 1) = 2.5
    trajectoryPoints(5, 2) = 3#    'above target
    
    trajectoryPoints(6, 1) = 3.5
    trajectoryPoints(6, 2) = 2.5   'above target
    
    trajectoryPoints(7, 1) = 4#
    trajectoryPoints(7, 2) = 2.5   'above target
    
    trajectoryPoints(8, 1) = 4.5
    trajectoryPoints(8, 2) = 2.1    'above target
    
    trajectoryPoints(9, 1) = 5#
    trajectoryPoints(9, 2) = 1.5    'inside target third crossing
    
    trajectoryPoints(10, 1) = 5.5
    trajectoryPoints(10, 2) = -0.5  'below target fourth crossing
    
    trajectoryPoints(11, 1) = 6#
    trajectoryPoints(11, 2) = -1.5  'below target
    
    
    'Act:
    Dim tracker As CTrajectoryTracker
    Set tracker = TrajectoryTracker(nPoints, yMin, yMax)
    Dim i As Long
    For i = LBound(trajectoryPoints, 1) To UBound(trajectoryPoints, 1)
        tracker.Track trajectoryPoints(i, 1), trajectoryPoints(i, 2)
    Next i
    
    Dim trackingList As New ArrayList
    tracker.FindCrossings trackingList
    
    Dim pbrDataList As New ArrayList
    tracker.BuildPointBlankRangeList trackingList, pbrDataList
    
    'Assert:
    Assert.AreEqual nCrossingPoints, trackingList.Count
    Assert.AreEqual nPBRIntervals, pbrDataList.Count
   

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


