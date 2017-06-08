---
title: Application.TimescaleNonWorkingEx Method (Project)
keywords: vbapj.chm2151
f1_keywords:
- vbapj.chm2151
ms.prod: project-server
api_name:
- Project.Application.TimescaleNonWorkingEx
ms.assetid: 50c1b96a-a91c-d538-07b7-44b048c8052b
ms.date: 06/08/2017
---


# Application.TimescaleNonWorkingEx Method (Project)

Sets the format of nonworking times, where color values can be hexadecimal RGB values.


## Syntax

 _expression_. **TimescaleNonWorkingEx**( ** _Draw_**, ** _Calendar_**, ** _Color_**, ** _Pattern_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Draw_|Optional|**Integer**|How nonworking times are denoted in relation to Gantt bars. Can be one of the following  **[PjNonWorkingPlacement](pjnonworkingplacement-enumeration-project.md)** constants: **pjBehind**, **pjInFront**, or **pjDoNotDraw**.|
| _Calendar_|Optional|**String**|The name of the calendar to format.|
| _Color_|Optional|**Long**|The color of nonworking times. Can be a hexadecimal value for the RGB color, where red is the last byte. For example, the value &;HFF0000 is blue and &;H00FFFF is yellow. |
| _Pattern_|Optional|**Integer**|The pattern for nonworking times. Can be one of the  **[PjFillPattern](pjfillpattern-enumeration-project.md)** constants.|

### Return Value

 **Boolean**


## Remarks

Using the  **TimescaleNonWorkingEx** method without specifying any arguments displays the **Timescale** dialog box with the **Non-working Time** tab selected.


## Example

The following example draws nonworking time behind the task bars in a light red.


```vb
Sub Timescale_NonWorking() 
    ' Sets nonworking time behind the task bars to red. 
 
    'Activate Gantt Chart. 
    ViewApply Name:="&;Gantt Chart" 
    TimescaleNonWorkingEx Draw:=pjBehind, Color:=&;HAAAAFF 
End Sub
```


