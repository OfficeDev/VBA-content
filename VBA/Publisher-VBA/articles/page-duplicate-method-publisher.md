---
title: Page.Duplicate Method (Publisher)
keywords: vbapb10.chm393256
f1_keywords:
- vbapb10.chm393256
ms.prod: publisher
api_name:
- Publisher.Page.Duplicate
ms.assetid: 9ef9d493-d2ca-8cac-3cce-6f0878acb288
ms.date: 06/08/2017
---


# Page.Duplicate Method (Publisher)

Creates a duplicate of the specified  **Page** object and then returns the new **Page** object.


## Syntax

 _expression_. **Duplicate**

 _expression_A variable that represents a  **Page** object.


### Return Value

Page


## Example

The following example duplicates the first page in the publication and then sets properties for the duplicate. A shape is then added to the new page and properties are set for the shape.


```vb
Dim objPage As Page 
Set objPage = ActiveDocument.Pages(1).Duplicate 
With objPage 
 .Background.Fill.ForeColor.SchemeColor = pbSchemeColorAccent1 
 .Shapes.AddShape msoShapeRectangle, 150, 250, 310, 275 
 With .Shapes(1) 
 .Fill.ForeColor.SchemeColor = pbSchemeColorAccent3 
 End With 
End With 

```


