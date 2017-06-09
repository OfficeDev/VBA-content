---
title: LayoutGuides Object (Publisher)
keywords: vbapb10.chm1179647
f1_keywords:
- vbapb10.chm1179647
ms.prod: publisher
api_name:
- Publisher.LayoutGuides
ms.assetid: 7430c1c4-c7f5-d9b6-cea8-b21fe9e2905f
ms.date: 06/08/2017
---


# LayoutGuides Object (Publisher)

Represents the measurement grid that appears superimposed on publication pages as an aid to laying out design elements.
 


## Example

Use the  **[LayoutGuides](document-layoutguides-property-publisher.md)** property of the **Document** object to return a **LayoutGuides** object. Use the **LayoutGuide** object's margin properties and **Rows** and **Columns** properties to set how many rows and columns are displayed in the layout guides and where they appear on a page.
 

 

 

 
This example sets the margins of the active presentation to two inches.
 

 



```
With ActiveDocument.LayoutGuides 
 .MarginTop = Application.InchesToPoints(Value:=2) 
 .MarginBottom = Application.InchesToPoints(Value:=2) 
 .MarginLeft = Application.InchesToPoints(Value:=2) 
 .MarginRight = Application.InchesToPoints(Value:=2) 
End With
```


## Properties



|**Name**|
|:-----|
|[Application](layoutguides-application-property-publisher.md)|
|[ColumnGutterWidth](layoutguides-columngutterwidth-property-publisher.md)|
|[Columns](layoutguides-columns-property-publisher.md)|
|[GutterCenterlines](layoutguides-guttercenterlines-property-publisher.md)|
|[HorizontalBaseLineOffset](layoutguides-horizontalbaselineoffset-property-publisher.md)|
|[HorizontalBaseLineSpacing](layoutguides-horizontalbaselinespacing-property-publisher.md)|
|[MarginBottom](layoutguides-marginbottom-property-publisher.md)|
|[MarginLeft](layoutguides-marginleft-property-publisher.md)|
|[MarginRight](layoutguides-marginright-property-publisher.md)|
|[MarginTop](layoutguides-margintop-property-publisher.md)|
|[MirrorGuides](layoutguides-mirrorguides-property-publisher.md)|
|[Parent](layoutguides-parent-property-publisher.md)|
|[RowGutterWidth](layoutguides-rowgutterwidth-property-publisher.md)|
|[Rows](layoutguides-rows-property-publisher.md)|
|[VerticalBaseLineOffset](layoutguides-verticalbaselineoffset-property-publisher.md)|
|[VerticalBaseLineSpacing](layoutguides-verticalbaselinespacing-property-publisher.md)|

