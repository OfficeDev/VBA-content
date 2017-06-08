---
title: Indexes Object (Word)
ms.prod: word
ms.assetid: 0441446a-c1b5-d333-5950-906fe463b61d
ms.date: 06/08/2017
---


# Indexes Object (Word)

A collection of  **[Index](index-object-word.md)** objects that represents all the indexes in the specified document.


## Remarks

Use the  **Indexes** property to return the **Indexes** collection. The following example formats indexes in the active document with the classic format.


```vb
ActiveDocument.Indexes.Format = wdIndexClassic
```

Use the  **Add** method to create an index and add it to the **Indexes** collection. The following example creates an index at the end of the active document.




```vb
Set myRange = ActiveDocument.Content 
myRange.Collapse Direction:=wdCollapseEnd 
ActiveDocument.Indexes.Add Range:=myRange, Type:=wdIndexRunin
```

Use  **Indexes** (Index), where Index is the index number, to return a single **Index** object. The index number represents the position of the **Index** object in the document. The following example updates the first index in the active document.




```vb
If ActiveDocument.Indexes.Count >= 1 Then 
 ActiveDocument.Indexes(1).Update 
End If
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

