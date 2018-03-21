VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimeEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class module for keeping information of single Redmine Time Entry
Private m_sUniqueKey As String ' Unique Key for sorting
Private m_sIssueID As String ' Redmine Issue ID
Private m_sIssueSubject As String ' Redmine Issue Subject
Private m_sIssueTracker As String ' Redmine Issue Tracker
Private m_sResource As String ' Redmine Issue Resource
Private m_sWorkType As String ' Work type of Redmine Time Entry
Private m_sComment As String ' Comment of Redmine Time Entry
Private m_sSpent As String ' Spent on of Redmine Time Entry
Private m_sHours As Double '  Hours of Redmine Time Entry
Private m_sOvertime As Boolean ' overtime hours of Redmine Time Entry. It requires specific parameter in Redmine configuration

Property Get UniqueKey() As String
    UniqueKey = m_sUniqueKey
End Property

Property Let UniqueKey(ByVal sUniqueKey As String)
    m_sUniqueKey = sUniqueKey
End Property

Property Get IssueId() As String
    IssueId = m_sIssueID
End Property

Property Let IssueId(ByVal sIssueId As String)
    m_sIssueID = sIssueId
End Property

Property Get IssueSubject() As String
    IssueSubject = m_sIssueSubject
End Property

Property Let IssueSubject(ByVal sIssueSubject As String)
    m_sIssueSubject = sIssueSubject
End Property

Property Get IssueTracker() As String
    IssueTracker = m_sIssueTracker
End Property

Property Let IssueTracker(ByVal sIssueTracker As String)
    m_sIssueTracker = sIssueTracker
End Property

Property Get Resource() As String
    Resource = m_sResource
End Property

Property Let Resource(ByVal sResource As String)
    m_sResource = sResource
End Property

Property Get WorkType() As String
    WorkType = m_sWorkType
End Property

Property Let WorkType(ByVal sWorkType As String)
    m_sWorkType = sWorkType
End Property

Property Get Comment() As String
    Comment = m_sComment
End Property

Property Let Comment(ByVal sComment As String)
    m_sComment = sComment
End Property

Property Get Spent() As String
    Spent = m_sSpent
End Property

Property Let Spent(ByVal sSpent As String)
    m_sSpent = sSpent
End Property

Property Get Hours() As Double
    Hours = m_sHours
End Property

Property Let Hours(ByVal sHours As Double)
    m_sHours = sHours
End Property

Property Get OverTime() As Boolean
    OverTime = m_sOvertime
End Property

Property Let OverTime(ByVal sOvertime As Boolean)
    m_sOvertime = sOvertime
End Property
