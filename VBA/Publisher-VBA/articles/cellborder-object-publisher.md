---
title: CellBorder Object (Publisher)
keywords: vbapb10.chm5308415
f1_keywords:
- vbapb10.chm5308415
ms.prod: publisher
api_name:
- Publisher.CellBorder
ms.assetid: c4eddeac-54cd-95ff-9423-b06e515a720e
ms.date: 06/08/2017
---


# CellBorder Object (Publisher)

Represents the color and weight settings for cell borders.
 


## Example

Use the various border properties of the  **Cell** object to return the different borders of a cell (left, right, top, bottom, and diagonal). The following example retrieves the top border of the first cell in a table.
 

 

```
Dim cbTemp As CellBorder 
 
Set cbTemp = ActiveDocument.Pages(1) _ 
 .Shapes(1).Table.Cells.Item(1).BorderTop
```

Use the  **[Color](cellborder-color-property-publisher.md)** and **[Weight](cellborder-weight-property-publisher.md)** properties of the **CellBorder** object to format the appearance of a cell border. The following example makes the left border of the first cell in a table red and two points thick.
 

 



```
Dim cbTemp As CellBorder 
 
Set cbTemp = ActiveDocument.Pages(1) _ 
 .Shapes(1).Table.Cells.Item(1).BorderLeft 
 
cbTemp.Color.RGB = RGB(255, 0, 0) 
cbTemp.Weight = 2
```


## Properties



|**Name**|
|:-----|
|[Application](cellborder-application-property-publisher.md)|
|[Color](cellborder-color-property-publisher.md)|
|[Parent](cellborder-parent-property-publisher.md)|
|[Weight](cellborder-weight-property-publisher.md)|

