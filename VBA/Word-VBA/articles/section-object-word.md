---
title: Section Object (Word)
keywords: vbawd10.chm2393
f1_keywords:
- vbawd10.chm2393
ms.prod: word
api_name:
- Word.Section
ms.assetid: 3fe563d8-fc05-c17a-e67b-c50eea7e7f13
ms.date: 06/08/2017
---


# Section Object (Word)

Represents a single section in a selection, range, or document. The  **Section** object is a member of the **[Sections](sections-object-word.md)** collection. The **Sections** collection includes all the sections in a selection, range, or document.


## Remarks

Use  **Sections** (Index), where Index is the index number, to return a single **Section** object. The following example changes the left and right page margins for the first section in the active document.


```
With ActiveDocument.Sections(1).PageSetup 
 .LeftMargin = InchesToPoints(0.5) 
 .RightMargin = InchesToPoints(0.5) 
End With
```

Use the  **Add** method or the **InsertBreak** method to add a new section to a document. The following example adds a new section at the beginning of the active document.




```
Set myRange = ActiveDocument.Range(Start:=0, End:=0) 
ActiveDocument.Sections.Add Range:=myRange 
myRange.InsertParagraphAfter
```

The following example adds a section break above the first paragraph in the selection.




```
Selection.Paragraphs(1).Range.InsertBreak _ 
 Type:=wdSectionBreakContinuous
```


 **Note**  The  **Headers** and **Footers** properties of the specified **Section** object return a **HeadersFooters** object.


## Properties



|**Name**|
|:-----|
|[Application](section-application-property-word.md)|
|[Borders](section-borders-property-word.md)|
|[Creator](section-creator-property-word.md)|
|[Footers](section-footers-property-word.md)|
|[Headers](section-headers-property-word.md)|
|[Index](section-index-property-word.md)|
|[PageSetup](section-pagesetup-property-word.md)|
|[Parent](section-parent-property-word.md)|
|[ProtectedForForms](section-protectedforforms-property-word.md)|
|[Range](section-range-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
