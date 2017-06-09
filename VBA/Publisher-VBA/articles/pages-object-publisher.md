---
title: Pages Object (Publisher)
keywords: vbapb10.chm524287
f1_keywords:
- vbapb10.chm524287
ms.prod: publisher
api_name:
- Publisher.Pages
ms.assetid: d6b7262c-015c-dcf3-bff4-0091dd32b78f
ms.date: 06/08/2017
---


# Pages Object (Publisher)

Represents all the pages in a publication. The  **Pages** collection contains all the **[Page](page-object-publisher.md)** objects in a publication.
 


## Example

Use the  **[Add](pages-add-method-publisher.md)** method to add a new page to a publication. The following example adds a new page and a shape to the active publication.
 

 

```
Sub AddPageAndShape() 
 With ActiveDocument.Pages.Add(Count:=1, After:=1) 
 With .Shapes.AddShape(Type:=msoShape5pointStar, _ 
 Left:=72, Top:=72, Width:=50, Height:=50) 
 .Fill.ForeColor.RGB = RGB(Red:=128, Green:=50, Blue:=255) 
 .Line.ForeColor.RGB = RGB(Red:=75, Green:=50, Blue:=255) 
 End With 
 End With 
 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Add](pages-add-method-publisher.md)|
|[AddWizardPage](pages-addwizardpage-method-publisher.md)|
|[FindByPageID](pages-findbypageid-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](pages-application-property-publisher.md)|
|[Count](pages-count-property-publisher.md)|
|[Item](pages-item-property-publisher.md)|
|[Parent](pages-parent-property-publisher.md)|

