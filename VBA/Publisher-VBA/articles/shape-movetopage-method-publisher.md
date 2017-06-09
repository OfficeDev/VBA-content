---
title: Shape.MoveToPage Method (Publisher)
keywords: vbapb10.chm2228376
f1_keywords:
- vbapb10.chm2228376
ms.prod: publisher
api_name:
- Publisher.Shape.MoveToPage
ms.assetid: 1893035f-6739-7480-6ba0-2ca6a42355fa
ms.date: 06/08/2017
---


# Shape.MoveToPage Method (Publisher)

Moves a shape to the specified page.


## Syntax

 _expression_. **MoveToPage**( **_Page_**,  **_Left_**,  **_Top_**)

 _expression_A variable that represents a  **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Page|Required| **Long**|Page to which the shape should be moved.|
|Left|Optional| **Variant**|Left position of the shape on the page.|
|Top|Optional| **Variant**|Top position of the shape on the page.|

## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **MoveToPage** method to move the first shape in the **Shapes** collection on the first page of a publication to the same relative location on the second page of the publication.

This code assumes that the current publication contains at least two pages, and that there is at least one shape on the first page of the publication.




```vb
Public Sub MoveToPage_Example() 
 
 Dim pubShape As Publisher.Shape 
 
 Set pubShape = ThisDocument.Pages(1).Shapes(1) 
 
 pubShape.MoveToPage 2 
 
End Sub
```


