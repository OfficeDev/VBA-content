---
title: Application.ZoomTimescale Method (Project)
keywords: vbapj.chm307
f1_keywords:
- vbapj.chm307
ms.prod: project-server
api_name:
- Project.Application.ZoomTimescale
ms.assetid: d20b2c8a-bef2-5456-73f1-a6fa417b427e
ms.date: 06/08/2017
---


# Application.ZoomTimescale Method (Project)

Zooms in on or out from the Gantt Chart, Resource Graph, Resource Usage, or Task Usage view to show information about tasks or resources in a certain duration.


## Syntax

 _expression_. **ZoomTimescale**( ** _Duration_**, ** _Entire_**, ** _Selection_**, ** _Reset_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Duration_|Optional|**Variant**|The duration to display in the view.|
| _Entire_|Optional|**Boolean**|**True** if the view resizes to fit the entire project onto the screen. The default value is **False**.|
| _Selection_|Optional|**Boolean**|**True** if the view resizes to fit only the selected tasks onto the screen. The default value is **False**.|
| _Reset_|Optional|**Boolean**|**True** if the view is reset to its default size. The default value is **False**.|

### Return Value

 **Boolean**


## Remarks

Except for the Resource Graph, where the entire view is affected, all zooming occurs in the non-entry portion of the active view.


## Example

The following example attempts to fit the entire project onto the screen.


```vb
Sub Display() 
 ZoomTimescale Entire:=True 
End Sub
```


