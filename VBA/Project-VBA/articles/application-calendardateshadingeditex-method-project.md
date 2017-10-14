---
title: Application.CalendarDateShadingEditEx Method (Project)
keywords: vbapj.chm2147
f1_keywords:
- vbapj.chm2147
ms.prod: project-server
api_name:
- Project.Application.CalendarDateShadingEditEx
ms.assetid: 13382dff-e043-480e-a9f7-300d743bd62a
ms.date: 06/08/2017
---


# Application.CalendarDateShadingEditEx Method (Project)

Changes the background color and pattern of date boxes in the Calendar view.


## Syntax

 _expression_. **CalendarDateShadingEditEx**( ** _Item_**, ** _Pattern_**, ** _Color_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Item_|Required|**Long**|The type of calendar day to change. Can be one of the  **[PjCalendarShading](pjcalendarshading-enumeration-project.md)** constants.|
| _Pattern_|Optional|**Long**|The pattern for the type of date box specified by  **Item**. Can be one of the **[PjFillPattern](pjfillpattern-enumeration-project.md)** constants.|
| _Color_|Optional|**Long**|The color for the type of date box specified by  **Item**. Can be a hexadecimal value for the RGB color, where red is the last byte. For example, the value &;HFF0000 is blue and &;H00FFFF is yellow.|

### Return Value

 **Boolean**


## Remarks

Besides  _Item_, **CalendarDateShadingEditEx** requires either the _Pattern_ or _Color_ parameter, or both, to run without an error. For example, the following line in the **Immediate** pane of the VBE works correctly.


```vb
? CalendarDateShadingEditEx (PjCalendarShading.pjBaseWorking, , &;H01dddd)
```


## Example

The following example changes the background color of working days in the base calander to a stippled purple and the color of nonworking days to light gray.


```vb
Sub CalendarDate_ShadingEdit() 
    ' Activate the Caldender view. 
    ViewApply Name:="Calendar" 
 
    CalendarDateShadingEditEx Item:=pjBaseWorking, Pattern:=pjLightFillPattern, Color:=&;H900090 
    CalendarDateShadingEditEx Item:=pjBaseNonworking, Color:=&;HDDDDDD 
End Sub
```


 **Note**  If you use any of the  **PjColor** enumeration constants for the _Color_ parameter, the color will be nearly black. For example, the value of **pjGreen** is 9, which in the **CalendarDateBoxesEx** method is a very dark red. To use only the sixteen colors available with **PjColor** constants, use the **[CalendarDateShadingEdit](application-calendardateshadingedit-method-project.md)** method.


