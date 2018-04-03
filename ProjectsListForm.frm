VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProjectsListForm 
   Caption         =   "Select the project"
   ClientHeight    =   5100
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9735.001
   OleObjectBlob   =   "ProjectsListForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProjectsListForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    'ListBox1.Column = Array("Header1", "Header2", "Header3")
    ListBox1.ColumnWidths = "40;100;500"
End Sub

Private Sub UserForm_Activate()
    
    Dim MyProjects As Collection
    Dim i As Integer
    
    Set MyProjects = getProjects()
    For i = 0 To MyProjects.Count - 1
        ListBox1.AddItem
        ListBox1.List(i, 0) = MyProjects.item(i + 1)(0)
        ListBox1.List(i, 1) = MyProjects.item(i + 1)(1)
        ListBox1.List(i, 2) = MyProjects.item(i + 1)(2)
        If ListBox1.List(i, 0) = ActiveProject.ProjectSummaryTask.Text1 Then ListBox1.Selected(i) = True
    Next
    'ListBox1.ColumnHeads = True
    
End Sub

Private Sub CommandButton1_Click()
    setProjectID CStr(ListBox1.Value)
    Me.Hide
End Sub
