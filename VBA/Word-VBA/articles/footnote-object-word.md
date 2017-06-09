---
title: Footnote Object (Word)
keywords: vbawd10.chm2367
f1_keywords:
- vbawd10.chm2367
ms.prod: word
api_name:
- Word.Footnote
ms.assetid: 877340c4-14f9-4560-eaf8-2c6482a1ade8
ms.date: 06/08/2017
---


# Footnote Object (Word)

Represents a footnote positioned at the bottom of the page or beneath text. The  **Footnote** object is a member of the **Footnotes** collection. The **[Footnotes](footnotes-object-word.md)** collection represents the footnotes in a selection, range, or document.


## Remarks

Use  **Footnotes** (Index), where Index is the index number, to return a single **Footnote** object. The index number represents the position of the footnote in the selection, range, or document. The following example applies red formatting to the first footnote in the selection.


```
If Selection.Footnotes.Count >= 1 Then 
 Selection.Footnotes(1).Reference.Font.ColorIndex = wdRed 
End If
```

Use the  **Add** method to add a footnote to the **[Footnotes](footnotes-object-word.md)** collection. The following example inserts an automatically numbered footnote immediately after the selection.




```
Selection.Collapse Direction:=wdCollapseEnd 
ActiveDocument.Footnotes.Add Range:=Selection.Range , _ 
 Text:="The Willow Tree, (Lone Creek Press, 1996)."
```


 **Note**  Footnotes positioned at the end of a document or section are considered endnotes and are included in the  **[Endnotes](endnotes-object-word.md)** collection.


## Methods



|**Name**|
|:-----|
|[Delete](footnote-delete-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Application](footnote-application-property-word.md)|
|[Creator](footnote-creator-property-word.md)|
|[Index](footnote-index-property-word.md)|
|[Parent](footnote-parent-property-word.md)|
|[Range](footnote-range-property-word.md)|
|[Reference](footnote-reference-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
