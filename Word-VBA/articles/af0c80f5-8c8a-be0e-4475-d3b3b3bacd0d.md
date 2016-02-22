
# Editor Object (Word)

Represents a single user who has been given specific permissions to edit portions of a document. 


## Remarks

Users who can be given permissions include individual contributors and groups of users as defined for Document Workspace sites.

The permissions you assign to ranges and selections go into effect only after a document is protected. Use the  **Editors** collection and the **Editor** object to assign specific permissions to sections of a document. Then use the **Protect** method to protect the document.

Use the  **Add** method of the **Editors** collection to give a specified user or group permission to modify a range or selection within a document. The following example gives the current user editing permission to modify the active selection.




```vb
Dim objEditor As Editor 
 
Set objEditor = Selection.Editors.Add(wdEditorCurrent)
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
[Editor Object Members](d7c78e7a-b04d-a6d4-4115-f4502d819b0b.md)
