---
title: Application.DocMove Method (Project)
keywords: vbapj.chm2015
f1_keywords:
- vbapj.chm2015
ms.prod: project-server
api_name:
- Project.Application.DocMove
ms.assetid: defa6ea7-5d1a-d3c4-6486-39192d1da99c
ms.date: 06/08/2017
---


# Application.DocMove Method (Project)

Moves the active window within the application window.


## Syntax

 _expression_. **DocMove**( ** _XPosition_**, ** _YPosition_**, ** _Points_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _XPosition_|Optional|**Long**| A number that specifies the distance of the active window from the left edge of the application.|
| _YPosition_|Optional|**Long**| A number that specifies the distance of the active window from the top edge of the application.|
| _Points_|Optional|**Boolean**|**True** if **XPosition** and **YPosition** are measured in points. **False** if they are measured in pixels. The default value is **False**.|

### Return Value

 **Boolean**


## Remarks

The positions specified are taken from the upper-left corner of the usable area of the application window. The usable area is the area remaining after removing the menu bar and toolbars. 


## Example

The following example moves the window of the active project to the upper-left corner of the main window.


```vb
Sub MoveProjectWindowToCorner() 
 DocMove 0, 0 
End Sub
```


