---
title: Application.AppMove Method (Project)
keywords: vbapj.chm2010
f1_keywords:
- vbapj.chm2010
ms.prod: project-server
api_name:
- Project.Application.AppMove
ms.assetid: 73ab96b7-4985-b25f-d202-89e6230e6e4e
ms.date: 06/08/2017
---


# Application.AppMove Method (Project)

Moves the main Project window.


## Syntax

 _expression_. **AppMove**( ** _XPosition_**, ** _YPosition_**, ** _Points_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _XPosition_|Optional|**Long**|A number that specifies the distance of the main window from the left edge of the screen.|
| _YPosition_|Optional|**Long**|A number that specifies the distance of the main window from the top edge of the screen.|
| _Points_|Optional|**Boolean**|**True** if **XPosition** and **YPosition** are measured in points. **False** if they are measured in pixels. The default value is **False**|

### Return Value

 **Boolean**


## Example

The following example moves the main window of Project nine points to the left.


```vb
Sub MoveMainWindowToLeft() 
    AppMove XPosition:=Application.Left - 9, Points:=True 
End Sub
```


