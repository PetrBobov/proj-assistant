VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Issue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Issue
' (c) Petr Bobov - https://github.com/PetrBobov/proj-assistant
'
'
' @author: petr.bobov@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Option Explicit

' Class module for keeping information of single Issue and activity if require
Private m_sID As String ' Issue ID
Private m_sSubject As String ' Issue Subject
Private m_sTracker As String ' Issue Tracker
Private m_sHyperlink As String 'Issue Hyperlink
Private m_sParentIssueID As String ' Issue Parent
Private m_sStartDate As String ' Issue start date
Private m_sDueDate As String ' Issue due date
Private m_sEstimatedHours As String ' Issue estimated hours
Private m_sActivity As String ' Activity if require

Property Get ID() As String
    ID = m_sID
End Property

Property Let ID(ByVal sID As String)
    m_sID = sID
End Property

Property Get Subject() As String
    Subject = m_sSubject
End Property

Property Let Subject(ByVal sSubject As String)
    m_sSubject = sSubject
End Property

Property Get Tracker() As String
    Tracker = m_sTracker
End Property

Property Let Tracker(ByVal sTracker As String)
    m_sTracker = sTracker
End Property

Property Get Hyperlink() As String
    Hyperlink = m_sHyperlink
End Property

Property Let Hyperlink(ByVal sHyperlink As String)
    m_sHyperlink = sHyperlink
End Property

Property Get ParentIssueID() As String
    ParentIssueID = m_sParentIssueID
End Property

Property Let ParentIssueID(ByVal sParentIssueID As String)
    m_sParentIssueID = sParentIssueID
End Property

Property Get StartDate() As String
    StartDate = m_sStartDate
End Property

Property Let StartDate(ByVal sStartDate As String)
    m_sStartDate = sStartDate
End Property

Property Get DueDate() As String
    DueDate = m_sDueDate
End Property

Property Let DueDate(ByVal sDueDate As String)
    m_sDueDate = sDueDate
End Property

Property Get EstimatedHours() As String
    EstimatedHours = m_sEstimatedHours
End Property

Property Let EstimatedHours(ByVal sEstimatedHours As String)
    m_sEstimatedHours = sEstimatedHours
End Property

Property Get Activity() As String
    Activity = m_sActivity
End Property

Property Let Activity(ByVal sActivity As String)
    m_sActivity = sActivity
End Property

Private Sub Class_Initialize()
    m_sID = ""
    m_sSubject = ""
    m_sTracker = ""
    m_sHyperlink = ""
    m_sParentIssueID = ""
    m_sStartDate = ""
    m_sDueDate = ""
    m_sEstimatedHours = ""
    m_sActivity = ""
End Sub
