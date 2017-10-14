---
title: RecentFiles Object (Word)
ms.prod: word
ms.assetid: c2d5e0b1-0d79-2fa7-c475-e5cace59ba1f
ms.date: 06/08/2017
---


# RecentFiles Object (Word)

A collection of  **[RecentFile](recentfile-object-word.md)** objects that represents the files that have been used recently. The items in the **RecentFiles** collection are displayed at the bottom of the **File** menu.


## Remarks

Use the  **RecentFiles** property to return the **RecentFiles** collection. The following example sets five as the maximum number of files that the **RecentFiles** collection can contain.


```
RecentFiles.Maximum = 5
```

Use the  **Add** method to add a file to the **RecentFiles** collection. The following example adds the active document to the list of recently-used files.




```vb
If ActiveDocument.Saved = True Then 
 RecentFiles.Add Document:=ActiveDocument.FullName, _ 
 ReadOnly:=True 
End If
```

Use  **RecentFiles** (Index), where Index is the index number, to return a single **RecentFile** object. The index number represents the position of the file on the **File** menu. The following example opens the first document in the **RecentFiles** collection.




```vb
If RecentFiles.Count >= 1 Then RecentFiles(1).Open
```

The  **SaveAs** and **Open** methods include an AddToRecentFiles argument that controls whether or not a file is added to the recently-used-files list when the file is opened or saved.


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

