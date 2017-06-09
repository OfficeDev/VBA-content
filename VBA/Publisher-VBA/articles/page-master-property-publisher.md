---
title: Page.Master Property (Publisher)
keywords: vbapb10.chm393222
f1_keywords:
- vbapb10.chm393222
ms.prod: publisher
api_name:
- Publisher.Page.Master
ms.assetid: f206b4f1-cde3-458d-f26c-a970ad3bd21b
ms.date: 06/08/2017
---


# Page.Master Property (Publisher)

Sets or returns a  **[Page](page-object-publisher.md)** object that represents the master page properties for the specified page.


## Syntax

 _expression_. **Master**

 _expression_A variable that represents a  **Page** object.


### Return Value

Page


## Remarks

Master pages do not have a  **Master** property. Any attempt to access the **Master** property of a master page will result in a run-time error.


## Example

This example adds a shape to the master page for the first page in the active publication.


```vb
Sub AddNewMasterPageShape() 
 With ActiveDocument.Pages(1).Master.Shapes.AddShape _ 
 (Type:=msoShape5pointStar, Left:=512, _ 
 Top:=50, Width:=50, Height:=50) 
 .Fill.ForeColor.CMYK.SetCMYK Cyan:=255, _ 
 Magenta:=255, Yellow:=0, Black:=0 
 End With 
End Sub
```

The  **Master** property can also be used to apply a master page to a page in a publication. The following example sets the master page of the first page of a publication to the master page of the second page in the publication. This example assumes that there are at least two pages and two master pages in the document.




```vb
ActiveDocument.Pages(1).Master = _ 
 ActiveDocument.Pages(2).Master
```


