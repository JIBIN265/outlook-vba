VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public WithEvents olInboxItems As Outlook.Items
Attribute olInboxItems.VB_VarHelpID = -1
Public WithEvents olApp As Outlook.Application
Attribute olApp.VB_VarHelpID = -1

Private Sub Application_Startup()
    Debug.Print "Application_Startup called"
    Set olApp = Outlook.Application
    InitializeInboxMonitor
End Sub

Public Sub InitializeInboxMonitor()
    Dim olNamespace As Outlook.NameSpace
    Dim olAccount As Outlook.Account
    Dim olInbox As Outlook.Folder
    
    Set olNamespace = Application.GetNamespace("MAPI")
    
    ' Find the specific account
    For Each olAccount In olNamespace.Accounts
        If olAccount.SmtpAddress = "jibin.thomas26@outlook.com" Then
            Set olInbox = olAccount.DeliveryStore.GetDefaultFolder(olFolderInbox)
            Set olInboxItems = olInbox.Items
            Debug.Print "Inbox monitor initialized for " & olAccount.SmtpAddress
            Exit For
        End If
    Next olAccount
    
    If olInboxItems Is Nothing Then
        Debug.Print "Could not find the specified email account"
    End If
End Sub

Private Sub olInboxItems_ItemAdd(ByVal item As Object)
    On Error Resume Next
    
    Debug.Print "olInboxItems_ItemAdd triggered"
    
    If TypeName(item) = "MailItem" Then
        Dim olMail As Outlook.mailItem
        Set olMail = item
        
        ProcessEmail olMail
    Else
        Debug.Print "Item is not a MailItem"
    End If
    
    On Error GoTo 0
End Sub

Private Sub ProcessEmail(olMail As Outlook.mailItem)
    Debug.Print "Processing email: " & olMail.Subject
    
    If InStr(1, olMail.Subject, "Process Invoice", vbTextCompare) > 0 Then
        Debug.Print "Subject contains 'Process Invoice'"
        Module1.CallSAPProcessAutomationAPI  ' Call the function from Module1
    Else
        Debug.Print "Subject does not contain 'Process Invoice'"
        Debug.Print "Full subject: '" & olMail.Subject & "'"
    End If
End Sub

Public Sub RebuildEventConnection()
    Set olInboxItems = Nothing
    InitializeInboxMonitor
End Sub

Public Sub TestNewMailEvent()
    Dim olNamespace As Outlook.NameSpace
    Dim olAccount As Outlook.Account
    Dim testMail As mailItem
    
    Set olNamespace = Application.GetNamespace("MAPI")
    
    ' Find the specific account
    For Each olAccount In olNamespace.Accounts
        If olAccount.SmtpAddress = "jibin.thomas26@outlook.com" Then
            Set testMail = olAccount.DeliveryStore.GetDefaultFolder(olFolderDrafts).Items.Add
            Exit For
        End If
    Next olAccount
    
    If testMail Is Nothing Then
        Debug.Print "Could not find the specified email account"
        Exit Sub
    End If
    
    testMail.Subject = "Test: Process Invoice"
    testMail.Save
    Debug.Print "Test email created in " & olAccount.SmtpAddress
    
    ProcessEmail testMail
End Sub
