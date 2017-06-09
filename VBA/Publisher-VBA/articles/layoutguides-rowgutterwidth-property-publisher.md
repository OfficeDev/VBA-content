---
title: LayoutGuides.RowGutterWidth Property (Publisher)
keywords: vbapb10.chm1114129
f1_keywords:
- vbapb10.chm1114129
ms.prod: publisher
api_name:
- Publisher.LayoutGuides.RowGutterWidth
ms.assetid: a7629683-68d2-4953-4c95-7e79e431f9c4
ms.date: 06/08/2017
---


# LayoutGuides.RowGutterWidth Property (Publisher)

Returns or sets the width of the row gutters that are used by the  **LayoutGuides** object to aid in the process of laying out design elements. Read/write **Single**.


## Syntax

 _expression_. **RowGutterWidth**

 _expression_A variable that represents a  **LayoutGuides** object.


### Return Value

Single


## Remarks

The default width of row gutters is 0.4 inches.


## Example

The following example modifies the second master page of the active publication so that it has four rows and four columns, row gutter width of 0.75 inches, column gutter width of 0.5 inches, and center lines in the gutters. Any new pages added to the publication that use the second master page as a template will have these properties.


```vb
Dim theMasterPage As page 
Dim theLayoutGuides As LayoutGuides 
 
Set theMasterPage = ActiveDocument.MasterPages(2) 
Set theLayoutGuides = theMasterPage.LayoutGuides 
 
With theLayoutGuides 
 .Rows = 4 
 .Columns = 4 
 .RowGutterWidth = Application.InchesToPoints(0.75) 
 .ColumnGutterWidth = Application.InchesToPoints(0.5) 
 .GutterCenterlines = True 
End With
```


