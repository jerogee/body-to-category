''
' warningFromBodyToCategory
'
' Removes an offending tag phrase from the email body and sets a
' custom "external" category instead, so that email is still marked.
'
' @author j.geertzen@elsevier.com
' @license MIT (https://opensource.org/licenses/MIT/)
''
Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)
    Dim objNS As Outlook.NameSpace
    Dim objEM As Outlook.MailItem
    Dim objCat As Category
    Dim regex As Object
    Dim strIDs() As String
    Dim intX As Integer
    Dim catFound As Boolean

    ' Get namespace
    Set objNS = Application.GetNamespace("MAPI")

    ' The offending string to strip out (any * before and after included)
    Const s = "External email: use caution"

    ' Regex object
    Set regex = CreateObject("VBScript.RegExp")
    regex.MultiLine = True
    regex.IgnoreCase = True
    regex.Global = True

    ' Inspect all entities. If the string is found, assign an 'External'
    ' category and replace the string. Could be improved with a nice regex
    ' and newline/HTML stripping, but will do for now.
    strIDs = Split(EntryIDCollection, ",")
    For intX = 0 To UBound(strIDs)
        Set objEM = objNS.GetItemFromID(strIDs(intX))
        If InStr(1, objEM.Body, s) > 0 Then
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

            ' Set category to External
            objEM.Categories = "External"

            ' Convert any RTF first to HTML
            If objEM.BodyFormat = olFormatRichText Then
                objEM.BodyFormat = olFormatHTML
            End If

            ' Strip the offending string
            If objEM.BodyFormat = olFormatHTML Then
                regex.Pattern = "<p>.*?" & s & ".*?</p>[\r\n]*<p>&nbsp;.*?</p>"
                objEM.HTMLBody = regex.Replace(objEM.HTMLBody, "")
            Else
                regex.Pattern = "\*+\s*" & s & "\s*\*+"
                objEM.Body = regex.Replace(objEM.Body, "")
            End If

            ' Save changes
            objEM.Save
        End If
    Next

    ' Clean up
    Set objEM = Nothing
    Set objNS = Nothing
    Set objCat = Nothing
    Set regex = Nothing
End Sub