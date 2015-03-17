VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Application_Startup()

Dim Ns As Outlook.NameSpace
Dim startDate As Date
Dim workDay As Integer
Dim objTask As Outlook.TaskItem

Set Ns = Application.GetNamespace("MAPI")
Set TaskFolder = Ns.GetDefaultFolder(olFolderTasks).Items

workDay = 8 'In hours
startDate = Now
remindDate = DateAdd("h", workDay, startDate)

Set objTask = TaskFolder.Find("[Subject] ='Time to go home'")

If objTask Is Nothing Then
    Set objTask = Application.CreateItem(olTaskItem)
End If

If DateDiff("d", objTask.startDate, startDate) = 0 Then
    Exit Sub
End If

With objTask
    .Subject = "Time to go home"
    .startDate = remindDate
    .ReminderTime = remindDate
    .ReminderSet = True
    .Body = "You arrived at work at " & startDate
    .Save
End With

End Sub
