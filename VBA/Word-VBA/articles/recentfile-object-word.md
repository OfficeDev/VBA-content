---
title: RecentFile Object (Word)
keywords: vbawd10.chm2404
f1_keywords:
- vbawd10.chm2404
ms.prod: word
api_name:
- Word.RecentFile
ms.assetid: c8d7a06d-c340-2d35-d4a9-5d0cd4a07aab
ms.date: 06/08/2017
---


# RecentFile Object (Word)

Represents a recently used file. The  **RecentFile** object is a member of the **[RecentFiles](recentfiles-object-word.md)** collection.


## Remarks

The  **RecentFiles** collection includes all the files that have been used recently. The items in the **RecentFiles** collection are displayed at the bottom of the **File** menu.

Use  **RecentFiles** (Index), where Index is the index number, to return a single **RecentFile** object. The index number represents the position of the file on the **File** menu. The following example opens the first document in the **RecentFiles** collection.




```vb
If RecentFiles.Count >= 1 Then RecentFiles(1).Open
```

Use the  **Add** method to add a file to the **RecentFiles** collection. The following example adds the active document to the list of recently-used files.




```vb
If ActiveDocument.Saved = True Then 
 RecentFiles.Add Document:=ActiveDocument.FullName, _ 
 ReadOnly:=True 
End If
```

The  **SaveAs** and **Open** methods include an AddToRecentFiles argument that controls whether or not a file is added to the recently-used-files list when the file is opened or saved.


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


