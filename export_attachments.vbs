'This code is supposed to extract attachments in Lotus Notes if they are in a field.
'Agent code

Sub Initialize
Dim session As New NotesSession
Dim db As NotesDatabase
Dim doc As NotesDocument
Dim nextdoc as notesdocument
Dim dirpath As String
Dim view As NotesView
Dim fullpath As String
      
Set db = session.CurrentDatabase
Set view = db.GetView("AllDocumentsByForm")
Set doc = view.GetFirstDocument

 Dim num As Integer
 dirpath = "c:\Images"
 If (Dir$ (dirpath, 16) = "") Then
  Mkdir dirpath
End If

While not doc is nothing
set nextdoc = view.getnextdocument(db)
            
 '-- Loop through all attachments in document and detach to Notes Data Directory
      Dim rtitem As Variant     
      Set rtitem = doc.GetFirstItem( "attachments" )
                  
'-- if array of embedded objects exist then detach all attachments into the Notes Data directory
      If Isarray( rtitem.EmbeddedObjects ) Then
            Forall o In rtitem.EmbeddedObjects
                  If ( o.Type = EMBED_ATTACHMENT ) Then
                        fullpath = dirpath & "\" & o.source
                                    
                        Call o.ExtractFile( fullpath )
                                    
                  End If
            End Forall
        end if

set doc = nextdoc
Wend
End Sub