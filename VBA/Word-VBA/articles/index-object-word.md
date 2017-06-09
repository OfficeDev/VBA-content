---
title: Index Object (Word)
keywords: vbawd10.chm2429
f1_keywords:
- vbawd10.chm2429
ms.prod: word
api_name:
- Word.Index
ms.assetid: 6a2aab98-485b-01c3-8d9b-9e108b455e22
ms.date: 06/08/2017
---


# Index Object (Word)

Represents a single index. The  **Index** object is a member of the **Indexes** collection. The **[Indexes](indexes-object-word.md)** collection includes all the indexes in the specified document.


## Remarks

Use  **Indexes** (Index), where Index is the index number, to return a single **Index** object. The index number represents the position of the **Index** object in the document. The following example updates the first index in the active document.


```vb
If ActiveDocument.Indexes.Count >= 1 Then 
 ActiveDocument.Indexes(1).Update 
End If
```

Use the  **Add** method to create an index and add it to the **Indexes** collection. The following example creates an index at the end of the active document.




```vb
Set myRange = ActiveDocument.Content 
myRange.Collapse Direction:=wdCollapseEnd 
ActiveDocument.Indexes.Add Range:=myRange, Type:=wdIndexRunin
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


