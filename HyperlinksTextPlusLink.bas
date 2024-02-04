Sub HyperlinksGetAll()
'
' Hyperlinks
' Собрать все гиперссылки из документа
' и записать в новый документ "ссылка + адрес"
'
    Dim docCurrent As Document
    Dim hLink As Hyperlink
    Dim docNew As Document
    Set docCurrent = ActiveDocument
    Set docNew = Documents.Add
    Application.ScreenUpdating = False
    
    For Each hLink In docCurrent.HyperLinks
    'Debug.Print hLink.Address
    'Debug.Print hLink.TextToDisplay
    'Debug.Print hLink.Range & (" ::: ") & hLink.Address
    Set dest = docNew.Range 'destination
    dest.InsertAfter (hLink.Range & " ::: " & hLink.Address & vbCrLf)
    Next
    
    docNew.Activate
    Application.ScreenUpdating = True
    'Application.ScreenRefresh
End Sub
