---
title: Application.GetCellInfo Method (Project)
keywords: vbapj.chm131092
f1_keywords:
- vbapj.chm131092
ms.prod: project-server
api_name:
- Project.Application.GetCellInfo
ms.assetid: ddd531b1-e66d-5c70-c4ed-2e2b456e3a3b
ms.date: 06/08/2017
---


# Application.GetCellInfo Method (Project)

Gets the cell object at the specified coordinates.


## Syntax

 _expression_. **GetCellInfo**( ** _x_**, ** _y_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _x_|Required|**Long**|Horizontal coordinate on the grid.|
| _y_|Required|**Long**|Vertical coordinate on the grid.|

### Return Value

Cell


## Remarks

The coordinates x=0, y=0 specify the top-left corner of the grid. The coordinate value increases to the right for x coordinates and down for y coordinates. The value of x must be less than or equal to the number of columns in the view. The value of y must be less than or equal to the number of rows in the view.


## Example

The following example sets the cell at x=1, y=0 to red.


```vb
Dim c As Cell 
 Set c = Application.GetCellInfo(1, 0) 
 c.CellColor = pjRed 
```


