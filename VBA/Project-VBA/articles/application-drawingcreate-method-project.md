---
title: Application.DrawingCreate Method (Project)
keywords: vbapj.chm2306
f1_keywords:
- vbapj.chm2306
ms.prod: project-server
api_name:
- Project.Application.DrawingCreate
ms.assetid: fc146a90-8207-0708-4cca-2015912b284a
ms.date: 06/08/2017
---


# Application.DrawingCreate Method (Project)

Activates the drawing feature.


## Syntax

 _expression_. **DrawingCreate**( ** _Type_**, ** _Behind_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required|**Long**|The type of drawing to create. Can be one of the following  **PjShape** constants: **pjOLEObject**, **pjLine**, **pjArrow**, **pjRectangle**, **pjEllipse**, **pjArc**, **pjPolygon**, or **pjTextBox**.|
| _Behind_|Optional|**Boolean**|**True** if the drawing is created behind task bars. The default value is **False**.|

### Return Value

 **Boolean**


## Remarks

The  **DrawingCreate** method requires user interaction before additional code can be executed.


