---
title: Footnotes Object (Word)
ms.prod: word
ms.assetid: d46a0972-2784-4814-d547-30122a35cdc1
ms.date: 06/08/2017
---


# Footnotes Object (Word)

A collection of  **Footnote** objects that represent all the footnotes in a selection, range, or document.


## Remarks

Use the  **Footnotes** property to return the **Footnotes** collection. The following example changes all of the footnotes in the active document to endnotes.


```
ActiveDocument.Footnotes.SwapWithEndnotes
```

Use the  **Add** method to add a footnote to the **Footnotes** collection. The following example adds a footnote immediately after the selection.




```
Selection.Collapse Direction:=wdCollapseEnd 
ActiveDocument.Footnotes.Add Range:=Selection.Range , _ 
 Text:="The Willow Tree, (Lone Creek Press, 1996)."
```

Use  **Footnotes** (index), where index is the index number, to return a single **[Footnote](footnote-object-word.md)** object. The index number represents the position of the footnote in the selection, range, or document. The following example applies red formatting to the first footnote in the selection.




```
If Selection.Footnotes.Count >= 1 Then 
 Selection.Footnotes(1).Reference.Font.ColorIndex = wdRed 
End If
```


 **Note**  Footnotes positioned at the end of a document or section are considered endnotes and are included in the  **[Endnotes](endnotes-object-word.md)** collection.


## Methods



|**Name**|
|:-----|
|[Add](footnotes-add-method-word.md)|
|[Convert](footnotes-convert-method-word.md)|
|[Item](footnotes-item-method-word.md)|
|[ResetContinuationNotice](footnotes-resetcontinuationnotice-method-word.md)|
|[ResetContinuationSeparator](footnotes-resetcontinuationseparator-method-word.md)|
|[ResetSeparator](footnotes-resetseparator-method-word.md)|
|[SwapWithEndnotes](footnotes-swapwithendnotes-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Application](footnotes-application-property-word.md)|
|[ContinuationNotice](footnotes-continuationnotice-property-word.md)|
|[ContinuationSeparator](footnotes-continuationseparator-property-word.md)|
|[Count](footnotes-count-property-word.md)|
|[Creator](footnotes-creator-property-word.md)|
|[Location](footnotes-location-property-word.md)|
|[NumberingRule](footnotes-numberingrule-property-word.md)|
|[NumberStyle](footnotes-numberstyle-property-word.md)|
|[Parent](footnotes-parent-property-word.md)|
|[Separator](footnotes-separator-property-word.md)|
|[StartingNumber](footnotes-startingnumber-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
