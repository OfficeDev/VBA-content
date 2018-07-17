---
title: Application.TimescaleNonWorking Method (Project)
keywords: vbapj.chm914
f1_keywords:
- vbapj.chm914
ms.prod: project-server
api_name:
- Project.Application.TimescaleNonWorking
ms.assetid: bc43da1f-1854-d1ca-f44b-48f660f9336f
ms.date: 06/08/2017
---


# Application.TimescaleNonWorking Method (Project)

Sets the format of nonworking times.


## Syntax

 _expression_. **TimescaleNonWorking**( ** _Draw_**, ** _Calendar_**, ** _Color_**, ** _Pattern_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Draw_|Optional|**Integer**|How nonworking times are denoted in relation to Gantt bars. Can be one of the following  **[PjNonWorkingPlacement](pjnonworkingplacement-enumeration-project.md)** constants: **pjBehind**, **pjInFront**, or **pjDoNotDraw**.|
| _Calendar_|Optional|**String**|The name of the calendar to format.|
| _Color_|Optional|**Integer**|The color of nonworking times. Can be one of the  **[PjColor](pjcolor-enumeration-project.md)** constants.|
| _Pattern_|Optional|**Integer**|The pattern for nonworking times. Can be one of the  **[PjFillPattern](pjfillpattern-enumeration-project.md)** constants.|

### Return Value

 **Boolean**


## Remarks

Using the  **TimescaleNonWorking** method without specifying any arguments displays the **Timescale** dialog box with the **Non-working Time** tab selected.

To set nonworking time format by using a hexadecimal RGB value for color, see  **[TimescaleNonWorkingEx](application-timescalenonworkingex-method-project.md)**.


## Example

The following example draws nonworking time behind the task bars in red.


```vb
Sub Timescale_NonWorking() 
 'Sets nonworking time behind the task bars to red. 
 
 'Activate Gantt Chart. 
 ViewApply Name:="&;Gantt Chart" 
 TimescaleNonWorking Draw:=pjBehind, Color:=pjRed 
End Sub
```


