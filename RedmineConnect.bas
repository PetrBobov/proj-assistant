Attribute VB_Name = "RedmineConnect"
'
' RedmineConnect
' (c) Petr Bobov - https://github.com/PetrBobov/proj-assistant
'
'
' @author: petr.bobov@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Option Explicit

' common constants declaration
Const ContentType = "application/json" ' substitute
Const ResponseFormat = Json ' substitute
Const Insecure = True
' custom constants declaration
Const BaseURL = "http://200.200.200.200" ' base URL
Const KeyName = "X-Redmine-API-Key" ' redmine api key
Const KeyValue = "dc552a23892c23bhjadb3c639d49885a8e4e71b8" ' redmine api key value
Const UserIdFormat = "00" ' How much place may be in user id
Const IssueIdFormat = "0000" ' How much place may be in issue id
Const ActivityIdFormat = "00" ' How much place may be in activity id
Const TrackerID = "6" ' Default tracker ID
Const DaysDelta = -2 ' How much days shift back from today for default value inbox form
Const GetOvertime = True '  true if overtime values were marked in redmine with special boolean parameters which order was defined by OvertimePPos
Const OvertimePPos = 2 ' Position of overtime parameter in Redmine (Administration - Custom fields - Time Entries)

'-------------------

Private Function dateFrom() As String ' call dateFrom from user by user form, dateTo equal Now by default

    Dim sDateFrom As String

    sDateFrom = InputBox("Insert dateFrom to import actual work data (yyyy-mm-dd), example " + Format(VBA.DateAdd("d", DaysDelta, Now), "YYYY-MM-DD"), "Date From", Format(VBA.DateAdd("d", DaysDelta, Now), "YYYY-MM-DD"))

    If Not checkDateFormat(sDateFrom, "YYYY-MM-DD") Then
        MsgBox "dateFrom wrong value!"
        dateFrom = "-1"
    Else
        dateFrom = sDateFrom
    End If

End Function

Private Function checkDateFormat(sDate As String, sFormat As String) As Boolean
    
    If IsDate(sDate) Then
        If Format(sDate, sFormat) = sDate Then
            checkDateFormat = True
        Else
            checkDateFormat = False
        End If
    Else
        checkDateFormat = False
    End If

End Function

Public Sub SetVision()
    Application.ScreenUpdating = True
End Sub

Public Sub setActualWork()

Dim v_DateFrom As String ' dateFrom
Dim ActualWorks As Collection ' grouping and sorting Actual work from Redmine
Dim oActualWork As TimeEntry ' single object TimeEntry class
Dim oIssue As Issue ' exist Issue
Dim dateFrom As String ' date from which actual work request
Dim ProjectResources As Resources
Dim oResource As Resource ' single MS Project resource
Dim oTask As Task ' single MS Project task
Dim i, j As Integer ' counters
Dim ResAssigments As Assignments ' resource assigments
Dim oResAssigment As Assignment ' resource assigment
Dim NewTask As Task
Dim AssigmentExist As Boolean
Dim D As Date

' checking ProjectID in Project task Text1
If getProjectID = "" Then
    MsgBox "Must add ProjectID in Text1 of Project task"
    Exit Sub
End If

v_DateFrom = dateFrom
    
If v_DateFrom = "-1" Then Exit Sub

If MsgBox("Updating may be proceed long time. Are you sure?", vbExclamation + vbOKCancel, "Information") = vbCancel Then Exit Sub

'Application.ScreenUpdating = False

Set ActualWorks = GetTimes(v_DateFrom)

Set ProjectResources = ActiveProject.Resources

For Each oActualWork In ActualWorks
    For Each oResource In ProjectResources
        If oResource.Name = oActualWork.Resource Then
            Set ResAssigments = oResource.Assignments
            If ResAssigments.Count = 0 Then
                'Debug.Print "Create task"
                Set oIssue = New Issue
                Call getIssue(oActualWork.IssueId, oIssue)
                oIssue.Activity = oActualWork.WorkType
                Set NewTask = createTask(oIssue)
                Set oResAssigment = NewTask.Assignments.Add(ResourceID:=oResource.ID)
                Set NewTask = Nothing
                'Debug.Print "Insert time"
                Call setActualWorkk(oResAssigment, oActualWork)
            Else
                AssigmentExist = False
                For Each oResAssigment In ResAssigments
                    Set oTask = oResAssigment.Task
                    If (oTask.Text1 = oActualWork.IssueId) And (oTask.Text4 = oActualWork.WorkType) Then
                        AssigmentExist = True
                        Exit For
                    End If
                    Set oTask = Nothing
                Next
                
                If Not AssigmentExist Then
                    'Debug.Print "Create task"
                    Set oIssue = New Issue
                    Call getIssue(oActualWork.IssueId, oIssue)
                    oIssue.Activity = oActualWork.WorkType
                    Set NewTask = createTask(oIssue)
                    Set oResAssigment = NewTask.Assignments.Add(ResourceID:=oResource.ID)
                    Set NewTask = Nothing
                    Set oIssue = Nothing
                End If
                'Debug.Print "Insert time"
                Call setActualWorkk(oResAssigment, oActualWork)
                Set oResAssigment = Nothing
                Set oIssue = Nothing
                Set ResAssigments = Nothing
            End If
        End If
    Next
Next

'Application.ScreenUpdating = True

Call MsgBox("Updating finished", vbInformation, "Information")

End Sub

Private Function GetTimes(sDate As String) As Collection
    
    Dim RedmineClient As New WebClient
    Dim sGetTimes As Collection ' collection of TimeEntry
    Dim limit, offset As Integer
    RedmineClient.BaseURL = BaseURL
    RedmineClient.Insecure = Insecure

    ' Create a WebRequest for getting directions
    Dim TimeRequest As New WebRequest
    TimeRequest.Resource = "/time_entries.json"
    TimeRequest.Method = WebMethod.HttpGet

    ' Set the request format
    TimeRequest.SetHeader KeyName, KeyValue
    TimeRequest.ContentType = ContentType
    TimeRequest.ResponseFormat = ResponseFormat
    
    ' Add querystring to the request
    offset = 0
    limit = 100
    TimeRequest.AddQuerystringParam "project_id", getProjectID
    TimeRequest.AddQuerystringParam "spent_on", CStr("><" + sDate) + "|" + Format(VBA.DateAdd("d", -1, Now), "YYYY-MM-DD")

    Dim Response As WebResponse
    Set sGetTimes = New Collection
    
    Do
        TimeRequest.AddQuerystringParam "offset", CStr(offset)
        TimeRequest.AddQuerystringParam "limit", CStr(limit)
        
        ' Execute the request and work with the response
        Set Response = RedmineClient.Execute(TimeRequest)
        Call ProcessGetTimes(Response, sGetTimes)
        offset = offset + 100
        limit = limit + 100
    Loop While CInt(Response.Data("total_count")) > offset
    
    Set GetTimes = sGetTimes
    
    'Debug.Print CStr(RedmineClient.GetFullUrl(TimeRequest))
    
End Function

Private Sub ProcessGetTimes(Response As WebResponse, TimeEntries As Collection)
    
    Dim i As Integer ' counter
    Dim oTimeEntry As TimeEntry ' object TimeEntry class
    Dim item As TimeEntry ' exist object in collection
    Dim UKoTimeEntry As String ' unique key for grouping hours = UserID+IssueID+UserID+ActivityID+SpentOn
    
    If Response.StatusCode = WebStatusCode.Ok Then
        For i = 1 To Response.Data("time_entries").Count
            
            ' Sequence of parts of the key defines order used for task image on the gantt chart. This is the best sequence imho
            If GetOvertime Then
                UKoTimeEntry = CStr(Response.Data("time_entries")(i)("spent_on")) + _
                CStr(Format(Response.Data("time_entries")(i)("issue")("id"), IssueIdFormat)) + _
                CStr(Format(Response.Data("time_entries")(i)("user")("id"), UserIdFormat)) + _
                CStr(Format(Response.Data("time_entries")(i)("activity")("id"), ActivityIdFormat)) + _
                CStr(Response.Data("time_entries")(i)("custom_fields")(OvertimePPos)("value"))
            Else
                UKoTimeEntry = CStr(Response.Data("time_entries")(i)("spent_on")) + _
                CStr(Format(Response.Data("time_entries")(i)("issue")("id"), IssueIdFormat)) + _
                CStr(Format(Response.Data("time_entries")(i)("user")("id"), UserIdFormat)) + _
                CStr(Format(Response.Data("time_entries")(i)("activity")("id"), ActivityIdFormat))
            End If
            
            Set oTimeEntry = New TimeEntry
            
            oTimeEntry.UniqueKey = UKoTimeEntry
            oTimeEntry.IssueId = CStr(Response.Data("time_entries")(i)("issue")("id"))
            oTimeEntry.Resource = CStr(Response.Data("time_entries")(i)("user")("name"))
            oTimeEntry.WorkType = CStr(Response.Data("time_entries")(i)("activity")("name"))
            oTimeEntry.Spent = CStr(Response.Data("time_entries")(i)("spent_on"))
            
            If GetOvertime Then
                If Response.Data("time_entries")(i)("custom_fields")(OvertimePPos)("value") = 1 Then
                    oTimeEntry.OverTime = True
                Else
                    oTimeEntry.OverTime = False
                End If
            Else
                oTimeEntry.OverTime = False
            End If
            
            If KeyExistInCollection(TimeEntries, UKoTimeEntry) Then
                oTimeEntry.Comment = TimeEntries.item(UKoTimeEntry).Comment + VBA.vbCr + VBA.vbLf + CStr(Response.Data("time_entries")(i)("comments"))
                oTimeEntry.Hours = CDbl(TimeEntries.item(UKoTimeEntry).Hours) + CDbl(Response.Data("time_entries")(i)("hours"))
                TimeEntries.Remove (UKoTimeEntry)
            Else
                oTimeEntry.Comment = CStr(Response.Data("time_entries")(i)("comments"))
                oTimeEntry.Hours = CDbl(Response.Data("time_entries")(i)("hours"))
            End If
            
            If TimeEntries.Count = 0 Then
                TimeEntries.Add oTimeEntry, UKoTimeEntry
            Else
                For Each item In TimeEntries
                    If oTimeEntry.UniqueKey < item.UniqueKey Then
                        TimeEntries.Add oTimeEntry, UKoTimeEntry, item.UniqueKey
                        Exit For
                    End If
                Next
                If Not KeyExistInCollection(TimeEntries, UKoTimeEntry) Then TimeEntries.Add oTimeEntry, UKoTimeEntry
            End If
            
            Set oTimeEntry = Nothing
            
        Next
        
    Else
        Debug.Print "Error here: " & Response.StatusDescription
    End If

End Sub

Private Function KeyExistInCollection(coll As Collection, strKey As String) As Boolean
    Dim var As Object
    On Error Resume Next
    Set var = coll(strKey)
    KeyExistInCollection = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Function createTask(oIssue As Issue) As Task
    
    Dim ProjectTasks As Tasks
    Dim ParentTask As Task
    Dim sNewTask As Task
    Dim OutlineLevel As Integer
    
    Set ParentTask = searchParentTask(oIssue.ParentIssueID)
    
    If ParentTask Is Nothing Then
        Set ProjectTasks = ActiveProject.Tasks
        Set sNewTask = ProjectTasks.Add(oIssue.ID + ". " + oIssue.Subject)
        Set ProjectTasks = Nothing
        OutlineLevel = 1
    Else
        If TaskOutlineChildrenExist(ParentTask) Then
            Set sNewTask = ParentTask.OutlineChildren.Add(oIssue.ID + ". " + oIssue.Subject, ParentTask.ID + ParentTask.OutlineChildren.Count + 1)
            OutlineLevel = ParentTask.OutlineLevel + 1
        Else
            Set ProjectTasks = ActiveProject.Tasks
            Set sNewTask = ProjectTasks.Add(oIssue.ID + ". " + oIssue.Subject, ParentTask.ID + 1)
            OutlineLevel = ParentTask.OutlineLevel + 1
            Set ProjectTasks = Nothing
        End If
    End If

    sNewTask.Text1 = oIssue.ID
    sNewTask.Text2 = oIssue.Subject
    sNewTask.Text3 = oIssue.Tracker
    sNewTask.Text4 = oIssue.Activity
    sNewTask.HyperlinkAddress = oIssue.Hyperlink
    sNewTask.Text5 = oIssue.ParentIssueID
    sNewTask.OutlineLevel = OutlineLevel
    
    Set createTask = sNewTask
    Set sNewTask = Nothing
    Set ParentTask = Nothing
End Function

Private Sub getIssue(sIssueId As String, oIssue As Issue)
    
    Dim RedmineClient As New WebClient
    RedmineClient.BaseURL = BaseURL
    RedmineClient.Insecure = Insecure

    ' Create a WebRequest for getting directions
    Dim IssueRequest As New WebRequest
    IssueRequest.Resource = "/issues/{ID}.json"
    IssueRequest.Method = WebMethod.HttpGet
    IssueRequest.AddUrlSegment "ID", sIssueId
    

    ' Set the request format
    IssueRequest.SetHeader KeyName, KeyValue
    IssueRequest.ContentType = ContentType
    IssueRequest.ResponseFormat = ResponseFormat
    
    ' Execute the request and work with the response
    Dim Response As WebResponse
    
    Set Response = RedmineClient.Execute(IssueRequest)
    
    If Response.StatusCode = WebStatusCode.Ok Then
        oIssue.ID = CStr(Response.Data("issue")("id"))
        oIssue.Tracker = CStr(Response.Data("issue")("tracker")("name"))
        oIssue.Subject = CStr(Response.Data("issue")("subject"))
        oIssue.Hyperlink = Replace(RedmineClient.GetFullUrl(IssueRequest), ".json", "")
        'sStatus = CStr(Response.Data("issue")("status")("name"))
        
        Dim st As String
        st = CStr(Chr(34) & "parent" & Chr(34) & ":{")
        If InStr(1, Response.Content, st, vbTextCompare) <> 0 Then
            oIssue.ParentIssueID = CStr(Response.Data("issue")("parent")("id"))
        Else
            oIssue.ParentIssueID = ""
        End If
    Else
        oIssue.ID = ""
        'Debug.Print "Error here: " & Response.StatusDescription
    End If
End Sub

Private Function searchParentTask(sTaskID As String) As Task

Dim sTask As Task
Dim sID As String
Dim sSubject As String
Dim sTracker As String
Dim sWorkType As String
Dim sHyperlink As String
Dim ParentTaskID As String
Dim oIssue As Issue

    Set oIssue = New Issue
    
    ParentTaskID = 0
    For Each sTask In ActiveProject.Tasks
        If sTask.Text1 = sTaskID Then
            If sTask.Resources.Count = 0 Then
                Set searchParentTask = sTask
                Exit Function
            Else
                oIssue.ID = sTask.Text1
                oIssue.Subject = sTask.Text2
                oIssue.Tracker = sTask.Text3
                oIssue.Activity = sTask.Text4
                oIssue.Hyperlink = sTask.Hyperlink
                oIssue.ParentIssueID = sTask.Text5
                Set searchParentTask = Nothing
            End If
        End If
    Next
    
    If ParentTaskID = 0 Then
        Set searchParentTask = Nothing
    Else
        Set sTask = createTask(oIssue)
        Set searchParentTask = sTask
        Exit Function
    End If
    
    Set searchParentTask = Nothing
    Set oIssue = Nothing
    
End Function

Private Function TaskOutlineChildrenExist(sTask As Task) As Boolean

Dim OutlineChildren As Tasks
    
    Set OutlineChildren = sTask.OutlineChildren
    If OutlineChildren Is Nothing Then
        TaskOutlineChildrenExist = False
    Else
        If OutlineChildren.Count = 0 Then
            TaskOutlineChildrenExist = False
        Else
            TaskOutlineChildrenExist = True
        End If
    End If
    
    Set OutlineChildren = Nothing

End Function

Private Sub setActualWorkk(oResAssigment As Assignment, oTimeEntry As TimeEntry)

Dim TSVS As TimeScaleValues
Dim TSV As TimeScaleValue

    If oTimeEntry.OverTime Then
        Set TSVS = oResAssigment.TimeScaleData(CDate(oTimeEntry.Spent), CDate(oTimeEntry.Spent), pjAssignmentTimescaledActualOvertimeWork, pjTimescaleDays, 1)
        For Each TSV In TSVS
            TSV.Value = oTimeEntry.Hours * 60
        Next
        If Not (oTimeEntry.Comment = "") Then oResAssigment.AppendNotes (oTimeEntry.Spent + ": " + oTimeEntry.Comment + VBA.vbCrLf)
        Set TSVS = Nothing
    Else
        Set TSVS = oResAssigment.TimeScaleData(CDate(oTimeEntry.Spent), CDate(oTimeEntry.Spent), pjAssignmentTimescaledActualWork, pjTimescaleDays, 1)
        For Each TSV In TSVS
            TSV.Value = oTimeEntry.Hours * 60
        Next
        If Not (oTimeEntry.Comment = "") Then oResAssigment.AppendNotes (oTimeEntry.Spent + ": " + oTimeEntry.Comment + VBA.vbCrLf)
        Set TSVS = Nothing
    End If
    
    Set TSV = Nothing
    Set TSVS = Nothing

End Sub

Sub ProjectTasksToRedmineIssues() ' transferring project's tasks to redmine's issues by whether creating or updating

    Dim Tasks As Tasks
    Dim i As Integer
    Dim UpdateOk As Boolean
    Dim CreateOk As Boolean
    Dim ProjectID As String
    Dim oIssue As Issue
    '------------------------------'
    
    ProjectID = getProjectID
    Set Tasks = ActiveSelection.Tasks
    
    For i = 1 To Tasks.Count
        If Tasks(i).Text1 = "" Then ' creating new issue
            Set oIssue = New Issue
            If Tasks(i).OutlineLevel > 1 Then
                CreateOk = createRedmineIssue(ProjectID, Tasks(i), oIssue)
                If CreateOk Then
                    Tasks(i).Text1 = oIssue.ID
                    Tasks(i).Name = oIssue.ID + ". " + oIssue.Subject
                    Tasks(i).Text2 = oIssue.Subject
                    Tasks(i).Text3 = oIssue.Tracker
                    'Tasks(i).Text4 = oIssue.Activity
                    Tasks(i).Text5 = oIssue.ParentIssueID
                Else
                    Call MsgBox("Something wrong with creating issue " + CStr(Tasks(i).Name), vbCritical, "Warning")
                    Exit Sub
                End If
            Else
                If Tasks(i).OutlineLevel > 0 Then ' first level's tasks
                    CreateOk = createRedmineIssue(ProjectID, Tasks(i), oIssue)
                    If CreateOk Then
                        Tasks(i).Text1 = oIssue.ID
                        Tasks(i).Name = oIssue.ID + ". " + oIssue.Subject
                        Tasks(i).Text2 = oIssue.Subject
                        Tasks(i).Text3 = oIssue.Tracker
                        'Tasks(i).Text4 = oIssue.Activity
                        Tasks(i).Text5 = oIssue.ParentIssueID
                    Else
                        Call MsgBox("Something wrong with creating issue " + CStr(Tasks(i).Name), vbCritical, "Warning")
                        Exit Sub
                    End If
                End If
            End If
            Set oIssue = Nothing
        Else
            If Tasks(i).OutlineLevel > 1 Then
                UpdateOk = updateRedmineIssue(Tasks(i).Text1, Tasks(i))
                If UpdateOk Then
                    'Tasks(i).Text5 = ParentTask.Text1
                Else
                    Call MsgBox("Something wrong with updating issue #" + CStr(Tasks(i).Text1), vbCritical, "Warning")
                    Exit Sub
                End If
            Else
                If Tasks(i).OutlineLevel > 0 Then ' first level's tasks
                    UpdateOk = updateRedmineIssue(Tasks(i).Text1, Tasks(i))
                    If Not UpdateOk Then
                        Call MsgBox("Something wrong with updating issue #" + CStr(Tasks(i).Text1), vbCritical, "Warning")
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next
    
End Sub

Private Function createRedmineIssue(sProjectId As String, oTask As Task, oIssue As Issue) As Boolean
    
    Dim RedmineClient As New WebClient
    RedmineClient.BaseURL = BaseURL
    RedmineClient.Insecure = Insecure

    Dim CreateIssueRequest As New WebRequest
    CreateIssueRequest.Resource = "/issues.json"
    CreateIssueRequest.Method = WebMethod.HttpPost
    CreateIssueRequest.SetHeader KeyName, KeyValue
    CreateIssueRequest.ContentType = ContentType
    CreateIssueRequest.ResponseFormat = ResponseFormat
    
        Dim IssueBody As Dictionary
        Set IssueBody = New Dictionary
            IssueBody.Add "project_id", sProjectId
            IssueBody.Add "tracker_id", TrackerID
            IssueBody.Add "subject", oTask.Name
            
            Dim ParentTask As Task
            Set ParentTask = oTask.OutlineParent
            If Not (ParentTask.Text1 = "") Then
                IssueBody.Add "parent_issue_id", ParentTask.Text1
            End If
            Set ParentTask = Nothing
            
        Dim IssueRoot As Dictionary
        Set IssueRoot = New Dictionary
            IssueRoot.Add "issue", IssueBody
        
    Set CreateIssueRequest.Body = IssueRoot
    
    Dim Response As WebResponse
    Set Response = RedmineClient.Execute(CreateIssueRequest)

    If Response.StatusCode = WebStatusCode.Created Then
        Call getIssue(CStr(Response.Data("issue")("id")), oIssue)
        createRedmineIssue = True
    Else
        createRedmineIssue = False
        'Debug.Print "Error here: " & Response.StatusDescription
    End If

End Function

Function updateRedmineIssue(sIssueId As String, oTask As Task) As Boolean
    
    Dim oIssue As Issue
    Set oIssue = New Issue
    Call getIssue(sIssueId, oIssue)
    
    Dim RedmineClient As New WebClient
    RedmineClient.BaseURL = BaseURL
    RedmineClient.Insecure = Insecure

    Dim UpdateIssueRequest As New WebRequest
    UpdateIssueRequest.Resource = "/issues/{ID}.json"
    UpdateIssueRequest.Method = WebMethod.HttpPut
    UpdateIssueRequest.AddUrlSegment "ID", sIssueId
    UpdateIssueRequest.SetHeader KeyName, KeyValue
    UpdateIssueRequest.ContentType = ContentType
    UpdateIssueRequest.ResponseFormat = ResponseFormat

    Dim IssueBody As Dictionary
    Set IssueBody = New Dictionary
        If Not (oTask.Text2 = "") And Not (oTask.Text2 = oIssue.Subject) Then
            IssueBody.Add "subject", oTask.Text2
        End If
        
        Dim ParentTask As Task
        Set ParentTask = oTask.OutlineParent
        If Not (ParentTask.Text1 = "") And Not (ParentTask.Text1 = oIssue.ParentIssueID) Then
            IssueBody.Add "parent_issue_id", ParentTask.Text1
        End If
        Set ParentTask = Nothing
        
        If Not (oTask.Start = "") And Not (oTask.Start = oIssue.StartDate) Then
            IssueBody.Add "start_date", Format(oTask.Start, "YYYY-MM-DD")
        End If
        If Not (oTask.Finish = "") And Not (oTask.Finish = oIssue.DueDate) Then
            IssueBody.Add "due_date", Format(oTask.Finish, "YYYY-MM-DD")
        End If
        'If Not (oTask.BaselineWork = "") And Not (oTask.BaselineWork = oIssue.EstimatedHours) Then
        '    IssueBody.Add "estimated_hours", CStr(Round(oTask.BaselineWork / 60))
        'End If
        
        Dim IssueRoot As Dictionary
        Set IssueRoot = New Dictionary
            IssueRoot.Add "issue", IssueBody
    
    Set UpdateIssueRequest.Body = IssueRoot
    
    Dim Response As WebResponse
    Set Response = RedmineClient.Execute(UpdateIssueRequest)
    
    If Response.StatusCode = WebStatusCode.Ok Then
        Call getIssue(sIssueId, oIssue)
        ' to be deleted
        oTask.Name = oIssue.ID + ". " + oIssue.Subject
        oTask.Text1 = oIssue.ID
        oTask.Text2 = oIssue.Subject
        oTask.Text3 = oIssue.Tracker
        'oTask.Text4 = oIssue.Activity
        oTask.Text5 = oIssue.ParentIssueID
        'oTask.HyperlinkAddress = oIssue.Hyperlink
        updateRedmineIssue = True
    Else
        'Debug.Print "Error here: " & Response.StatusDescription
        updateRedmineIssue = False
    End If
    Set oIssue = Nothing
End Function

Private Function getProjectID()
    getProjectID = ActiveProject.ProjectSummaryTask.Text1
End Function

