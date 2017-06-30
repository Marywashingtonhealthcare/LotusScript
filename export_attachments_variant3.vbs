



Sub Initialize
    Dim Sess As New NotesSession
    Dim db As NotesDatabase
    Dim doc As NotesDocument
    Dim shortname As String
    Dim view As NotesView
    Dim rtitem As Variant

    Dim embedObj As NotesEmbeddedObject
    Dim Attachments As Variant
    Dim c as Double

    Set db = sess.CurrentDatabase
    Set view = db.GetView("YourViewGoesHere")
    Set doc = view.GetFirstDocument
    Set c = 1
    
    'The Loop where the magic happens
    Do Until doc Is Nothing

        If doc.HasEmbedded Then
            If Not(doc.YourNameField Is Nothing) Then
                shortname = doc.YourNameFieldHere(0)
            Else
                shortname = doc.TitleField(0) & "-NoFormNum"
            End If

            attachments = Evaluate("@AttachmentNames", doc)

            For x = 0 To UBound(attachments)
                Set embedObj = Doc.GetAttachment(CStr(attachments(x)))
                filepath = "C:\YourDirectoryHere\" & shortname & "-" & CStr(attachments(x))
                Call embedObj.Extractfile(filepath)
            Next
                    Call doc.Save( True, False )
        End If
           Set doc = view.GetNextDocument( doc )
        Print "Processing " & Cstr(c) & " documents."
        c = c + 1
     Loop
    
    Print "Processed " & Cstr(c) & " documents."
    
End Sub





