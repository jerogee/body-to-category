Attribute VB_Name = "BodyToCategory"
''
' @author j.geertzen@elsevier.com
' @license MIT (https://opensource.org/licenses/MIT/)
''

' Strip any instance of string s from the email body of MailItem objM,
' and if string s is found, assign a custom category ("External")
Public Sub BodyToCategory(ByRef objMI As Outlook.MailItem, ByVal s As String)
    Dim regex As Object

    ' Regex object
    Set regex = CreateObject("VBScript.RegExp")
    regex.MultiLine = True
    regex.IgnoreCase = True
    regex.Global = True

    ' If string found assign category and strip out
    If InStr(1, objMI.Body, s) > 0 Then
        ' Set category to External
        objMI.Categories = "External"

        ' Convert any RTF first to HTML
        If objMI.BodyFormat = olFormatRichText Then
            objMI.BodyFormat = olFormatHTML
        End If

        ' Strip the offending string
        If objMI.BodyFormat = olFormatHTML Then
            ' Direct pattern
            regex.Pattern = "<p><span[^>]*?><strong>\*+\s*" & s & "\s*\*+</strong></span></p>[\r\n]*<p>&nbsp;</p>"
            objMI.HTMLBody = regex.Replace(objMI.HTMLBody, "")

            ' Pattern after Outlook's HTML recoding (in case of Forward, re-send, etc.)
            regex.Pattern = "<p[^>]*?><strong[^>]*?><span[^>]*?>\*+\s*" & s & "\s*\*+</span></strong>.*?</p>"
            objMI.HTMLBody = regex.Replace(objMI.HTMLBody, "")
        Else
            regex.Pattern = "\*+\s*" & s & "\s*\*+"
            objMI.Body = regex.Replace(objMI.Body, "")
        End If

        ' Save changes
        objMI.Save
    End If

    ' Clean up
    Set regex = Nothing
End Sub

' Strip a specific tag string from email body
Public Sub StripTag(ByRef Item As Outlook.MailItem)
    ' The offending string to strip out (any * before and after included)
    Const s As String = "External email: use caution"

    ' Strip string frombody
    BodyToCategory Item, s
End Sub

' Check if external category has been defined. If not, define it.
Public Sub SetExternalCategory()
    Dim objNS As Outlook.NameSpace
    Dim objCat As Category
    Dim catFound As Boolean

    ' Get namespace
    Set objNS = Application.GetNamespace("MAPI")

    ' Check if External category has already been added. If not add.
    catFound = False
    For Each objCat In objNS.Categories
        If objCat.Name = "External" Then
            catFound = True
            Exit For
        End If
    Next
    If Not catFound Then
        objNS.Categories.Add "External", 2, 0
    End If

    ' Clean up
    Set objNS = Nothing
    Set objCat = Nothing
End Sub

' Process each MailItem in an EntryIDCollection, and strip out
' the offending string. Make sure that a custom "External" category
' exists. This needs to be called upon new email, in: Application_NewMailEx()
Public Sub StripCollection(EntryIDCollection As String)
    Dim objNS As Outlook.NameSpace
    Dim strIDs() As String
    Dim intX As Integer

    ' Set external category if needed
    SetExternalCategory

    ' Get namespace
    Set objNS = Application.GetNamespace("MAPI")

    ' Inspect all entities.
    strIDs = Split(EntryIDCollection, ",")
    For intX = 0 To UBound(strIDs)
        StripTag objNS.GetItemFromID(strIDs(intX))
    Next

    ' Clean up
    Set objNS = Nothing
End Sub

' Process each MailItem in the Inbox, and strip out
' the offending string. Make sure that a custom "External" category
' exists. This needs to be called upon startup, in: Application_Startup()
'
' Only process recent emails (last two days), otherwise takes too long.
' Call with do_all_messages = True to process all messages
Public Sub StripInbox(Optional do_all_messages As Boolean = False)
    Dim objNS As Outlook.NameSpace
    Dim objIB As Outlook.MAPIFolder
    Dim Items As Outlook.Items
    Dim Item As Object

    ' Set the number of most recent days to check
    ' (taken unless do_all_messages = True)
    Const n_last_days = 2

    ' Set external category if needed
    SetExternalCategory

    ' Get namespace and inbox
    Set objNS = Application.GetNamespace("MAPI")
    Set objIB = objNS.GetDefaultFolder(olFolderInbox)

    ' Inspect recent (n_last_days) email (or full inbox)
    Set Items = objIB.Items.Restrict("[ReceivedTime]>'" & Format(Date - n_last_days, "DDDDD HH:NN") & "'")
    If do_all_messages Then
        Set Items = objIB.Items
    End If
    For Each Item In Items
        If TypeOf Item Is Outlook.MailItem Then
            Debug.Print Item.SentOn & " - " & Item.Subject
            StripTag Item
        End If
    Next

    ' Clean up
    Set objNS = Nothing
    Set objIB = Nothing
    Set Items = Nothing
    Set Item = Nothing
End Sub
