---
title: Endnote Object (Word)
keywords: vbawd10.chm2366
f1_keywords:
- vbawd10.chm2366
ms.prod: word
api_name:
- Word.Endnote
ms.assetid: 01f29be4-58e7-28f5-5fcb-dae50c33890e
ms.date: 06/08/2017
---


# Endnote Object (Word)

Represents an endnote. The  **Endnote** object is a member of the **Endnotes** collection, which represents the endnotes in a selection, range, or document.


## Remarks

Use  **Endnotes** (Index), where Index is the index number, to return a single **Endnote** object. The index number represents the position of the endnote in the selection, range, or document. The following example applies red formatting to the first endnote in the selection.


```vb
If Selection.Endnotes.Count >= 1 Then 
 Selection.Endnotes(1).Reference.Font.ColorIndex = wdRed 
End If
```

Use the  **Add** method to add an endnote to the **[Endnotes](endnotes-object-word.md)** collection. The following example adds an endnote immediately after the selection.




```
Selection.Collapse Direction:=wdCollapseEnd 
ActiveDocument.Endnotes.Add Range:=Selection.Range , _ 
 Text:="The Willow Tree, (Lone Creek Press, 1996)."
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


