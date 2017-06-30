'This variant might be the one we want.



Sub Initialize
     Dim Sess As New NotesSession
     Dim db As NotesDatabase
     Dim doc As NotesDocument
     Dim shortname As String
     Dim DataDirectory As String
     Dim view As NotesView
     Dim rtitem As Variant
     Dim object As NotesEmbeddedObject
     
     Set db = sess.CurrentDatabase
        'pick view from database
    Set view = db.GetView("YourViewGoesHere")
     
     DataDirectory = Sess.GetEnvironmentString("\\nycweb04\TextBases\Reports\",True)
     
     Set doc = view.GetFirstDocument   
     
     Do Until doc Is Nothing
          Set rtitem= doc.GetFirstItem("PDFtext")
          If Not(rtitem Is Nothing) Then ' should check type as well...
                shortname = doc.IdNum(0)'This is the name of the field we want the file called
                 If doc.HasEmbedded Then
                       Set object= rtitem.EmbeddedObjects(0)
                                    'Replace independent with a subdirectory of the DATA directory
                       Call object.ExtractFile(DataDirectory & "\PDFs\" & shortname & ".pdf")
                
                    Call doc.Save( True, False )
                 End If
           End If
           Set doc = view.GetNextDocument( doc )
     Loop
End Sub



