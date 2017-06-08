---
title: Hyperlink Object (Word)
keywords: vbawd10.chm2461
f1_keywords:
- vbawd10.chm2461
ms.prod: word
api_name:
- Word.Hyperlink
ms.assetid: af785a9e-081a-e359-705f-04f490304e2e
ms.date: 06/08/2017
---


# Hyperlink Object (Word)

Represents a hyperlink. The  **Hyperlink** object is a member of the **Hyperlinks** collection.


## Remarks

Use the  **Hyperlink** property to return a **Hyperlink** object associated with a shape (a shape can have only one hyperlink). The following example activates the hyperlink associated with the first shape in the active document.


```
ActiveDocument.Shapes(1).Hyperlink.Follow
```

Use  **Hyperlinks** (Index), where Index is the index number, to return a single **Hyperlink** object from a document, range, or selection. The following example activates the first hyperlink in the selection.




```
If Selection.HyperLinks.Count >= 1 Then 
 Selection.HyperLinks(1).Follow 
End If
```


## Methods



|**Name**|
|:-----|
|[AddToFavorites](hyperlink-addtofavorites-method-word.md)|
|[CreateNewDocument](hyperlink-createnewdocument-method-word.md)|
|[Delete](hyperlink-delete-method-word.md)|
|[Follow](hyperlink-follow-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Address](hyperlink-address-property-word.md)|
|[Application](hyperlink-application-property-word.md)|
|[Creator](hyperlink-creator-property-word.md)|
|[EmailSubject](hyperlink-emailsubject-property-word.md)|
|[ExtraInfoRequired](hyperlink-extrainforequired-property-word.md)|
|[Name](hyperlink-name-property-word.md)|
|[Parent](hyperlink-parent-property-word.md)|
|[Range](hyperlink-range-property-word.md)|
|[ScreenTip](hyperlink-screentip-property-word.md)|
|[Shape](hyperlink-shape-property-word.md)|
|[SubAddress](hyperlink-subaddress-property-word.md)|
|[Target](hyperlink-target-property-word.md)|
|[TextToDisplay](hyperlink-texttodisplay-property-word.md)|
|[Type](hyperlink-type-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
