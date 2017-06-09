---
title: Hyperlink.PageID Property (Publisher)
keywords: vbapb10.chm4587525
f1_keywords:
- vbapb10.chm4587525
ms.prod: publisher
api_name:
- Publisher.Hyperlink.PageID
ms.assetid: 1b5051eb-e6b4-a5a7-610a-5be03863a92b
ms.date: 06/08/2017
---


# Hyperlink.PageID Property (Publisher)

Returns or sets a  **Long** indicating the page in the publication that is the destination for the specified hyperlink. Read/write.


## Syntax

 _expression_. **PageID**

 _expression_A variable that represents a  **Hyperlink** object.


## Example

The following example reports what page the first hyperlink in the active publication is linked to.


```vb
Dim hypTemp As Hyperlink 
Dim lngID As Long 
Dim strPage As String 
 
Set hypTemp = ActiveDocument.Pages(1).Shapes(1).Hyperlink 
 
lngID = hypTemp.PageID 
strPage = ActiveDocument.Pages.FindByPageID(PageID:=lngID).PageNumber 
 
MsgBox "This hyperlink goes to the page " &; strPage &; "."
```


