Attribute VB_Name = "Time"
Option Explicit

Public Sub TimeRefresh()
    Application.Calculate
    Application.OnTime Now + TimeValue("00:00:01"), "TimeRefresh"
End Sub
