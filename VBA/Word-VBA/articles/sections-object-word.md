---
title: Sections Object (Word)
keywords: vbawd10.chm2394
f1_keywords:
- vbawd10.chm2394
ms.prod: word
ms.assetid: cf6f77ba-9eee-5614-e697-bc031c4c6dcd
ms.date: 06/08/2017
---


# Sections Object (Word)

A collection of  **Section** objects in a selection, range, or document.


## Remarks

Use the  **Sections** property to return the **Sections** collection. The following example inserts text at the end of the last section in the active document.


```
With ActiveDocument.Sections.Last.Range 
 .Collapse Direction:=wdCollapseEnd 
 .InsertAfter "end of document" 
End With
```

Use the  **Add** method or the **InsertBreak** method to add a new section to a document. The following example adds a new section at the beginning of the active document.




```
Set myRange = ActiveDocument.Range(Start:=0, End:=0) 
ActiveDocument.Sections.Add Range:=myRange 
myRange.InsertParagraphAfter
```

The following example displays the number of sections in the active document, adds a section break above the first paragraph in the selection, and then displays the number of sections again.




```
MsgBox ActiveDocument.Sections.Count &amp; " sections" 
Selection.Paragraphs(1).Range.InsertBreak _ 
 Type:=wdSectionBreakContinuous 
MsgBox ActiveDocument.Sections.Count &amp; " sections"
```

Use  **Sections** (index), where index is the index number, to return a single **Section** object. The following example changes the left and right page margins for the first section in the active document.




```
With ActiveDocument.Sections(1).PageSetup 
 .LeftMargin = InchesToPoints(0.5) 
 .RightMargin = InchesToPoints(0.5) 
End With
```


## Methods



|**Name**|
|:-----|
|[Add](sections-add-method-word.md)|
|[Item](sections-item-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Application](sections-application-property-word.md)|
|[Count](sections-count-property-word.md)|
|[Creator](sections-creator-property-word.md)|
|[First](sections-first-property-word.md)|
|[Last](sections-last-property-word.md)|
|[PageSetup](sections-pagesetup-property-word.md)|
|[Parent](sections-parent-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
