---
title: Page.AddGuide Method (Visio)
keywords: vis_sdr.chm10916035
f1_keywords:
- vis_sdr.chm10916035
ms.prod: visio
api_name:
- Visio.Page.AddGuide
ms.assetid: 7be0cc07-6322-a3f0-3292-6dc66804db44
ms.date: 06/08/2017
---


# Page.AddGuide Method (Visio)

Adds a guide to a drawing page.


## Syntax

 _expression_ . **AddGuide**( **_Type_** , **_xPos_** , **_yPos_** )

 _expression_ A variable that represents a **Page** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **Integer**|The type of guide to add.|
| _xPos_|Required| **Double**|The x-coordinate of a point on the guide.|
| _yPos_|Required| **Double**|The y-coordinate of a point on the guide.|

### Return Value

Shape


## Remarks

The following constants declared by the Visio type library are valid values for guide types.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visPoint**|1|Guide point|
| **visHorz**|2|Horizontal guide|
| **visVert**|3|Vertical guide|

## Example

The following macro shows how to add a horizontal guide to a page.


```vb
 
Public Sub AddGuide_Example() 
 
 Dim vsoPages As Visio.Pages 
 Dim vsoPage As Visio.Page 
 Dim vsoShapes As Visio.Shapes 
 Dim vsoShape As Visio.Shape 
 Dim vsoPageHeightCell as Visio.Cell 
 Dim intPageHeightIU as Integer 
 
 'Get the Pages collection of the ThisDocument object. 
 Set vsoPages = ThisDocument.Pages 
 
 'Set the Page object to the first page of the Pages collection. 
 Set vsoPage = vsoPages(1) 
 
 'Get the Shapes collection of the vsoPage object. 
 Set vsoShapes = vsoPage.Shapes 
 
 'Get the page height in internal units. 
 Set vsoPageHeightCell = vsoPage.PageSheet.CellsSRC( _ 
 visSectionObject, visRowPage, visPageHeight) 
 intPageHeightIU = vsoPageHeightCell.ResultIU 
 
 'Add a guide to the Shapes collection and set it 
 'as the vsoShape object. The guide is a horizontal line 
 'running through the middle of the page. 
 Set vsoShape = vsoPage.AddGuide(visHorz,0,intPageHeightIU/2) 
 
End Sub
```


