---
title: Revision Object (Word)
keywords: vbawd10.chm2433
f1_keywords:
- vbawd10.chm2433
ms.prod: word
api_name:
- Word.Revision
ms.assetid: e6f64467-a438-88f1-60f9-975365a1430e
ms.date: 06/08/2017
---


# Revision Object (Word)

Represents a change marked with a revision mark. The  **Revision** object is a member of the **[Revisions](revisions-object-word.md)** collection. The **Revisions** collection includes all the revision marks in a range or document.


## Remarks

Use  **Revisions** (Index), where Index is the index number, to return a single **Revision** object. The index number represents the position of the revision in the range or document. The following example displays the author name for the first revision in section one of the active document.


```
MsgBox ActiveDocument.Sections(1).Range.Revisions(1).Author
```

The  **Add** method isn't available for the **Revisions** collection. **Revision** objects are added when change tracking is enabled. Set the **TrackRevisions** property to **True** to track revisions made to the document text. The following example enables revision tracking and then inserts "Action " before the selection.




```
ActiveDocument.TrackRevisions = True 
Selection.InsertBefore "Action "
```


## Methods



|**Name**|
|:-----|
|[Accept](revision-accept-method-word.md)|
|[Reject](revision-reject-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Application](revision-application-property-word.md)|
|[Author](revision-author-property-word.md)|
|[Cells](revision-cells-property-word.md)|
|[Creator](revision-creator-property-word.md)|
|[Date](revision-date-property-word.md)|
|[FormatDescription](revision-formatdescription-property-word.md)|
|[Index](revision-index-property-word.md)|
|[MovedRange](revision-movedrange-property-word.md)|
|[Parent](revision-parent-property-word.md)|
|[Range](revision-range-property-word.md)|
|[Style](revision-style-property-word.md)|
|[Type](revision-type-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
