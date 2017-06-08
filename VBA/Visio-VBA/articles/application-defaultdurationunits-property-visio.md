---
title: Application.DefaultDurationUnits Property (Visio)
keywords: vis_sdr.chm10051045
f1_keywords:
- vis_sdr.chm10051045
ms.prod: visio
api_name:
- Visio.Application.DefaultDurationUnits
ms.assetid: 11810de2-0c2f-a498-6b7a-090d5397066b
ms.date: 06/08/2017
---


# Application.DefaultDurationUnits Property (Visio)

Determines the default unit of measure for quantities that represent durations. Read/write.


## Syntax

 _expression_ . **DefaultDurationUnits**

 _expression_ A variable that represents an **Application** object.


### Return Value

Variant


## Remarks

The  **DefaultDurationUnits** property corresponds to the value shown in the **Duration** box under **Display** on the **Advanced** tab of the **Visio Options** dialog box (click the **File** tab, and then click **Options**).

The return value contains one of the values of  **[VisUnitCodes](visunitcodes-enumeration-visio.md)** , which are declared in the Microsoft Visio type library.

You can specify  **DefaultDurationUnits** as an integer (a member of **[VisUnitCodes](visunitcodes-enumeration-visio.md)** ) or a string value such as "minutes". If the string is invalid or the unit code is inappropriate (non-duration), an error is generated.

For a complete list of valid unit strings along with corresponding Automation constants (integer values), see [About Units of Measure](http://msdn.microsoft.com/library/b6140312-b8e6-0cf2-9fe0-b14e800216bf%28Office.15%29.aspx).

Cell formulas that contain a specific unit of measure are displayed in those units regardless of the default duration units setting. Many cell formulas, however, use implicit unit syntax and are displayed in default units.

A program can create a cell whose formula displays in default units by setting the cell's  **Formula** property to a string in implicit unit syntax. For example, if a formula specifying duration is "=10[em,E]" , the result displays as "0.0069 ed" if the **DefaultDurationUnits** property is **visElapsedDay** , and "600.0000 es" if the **DefaultDurationUnits** property is **visElapsedSec** .

Alternatively, a program can use the following statement to set the cell's result to default duration units: 




```
vsoCell.Result(visDurationUnits) = 60
```

In this case, the result is 60 minutes if the  **DefaultDurationUnits** property is **visElapsedMin** and 60 seconds if the **DefaultDurationUnits** property is **visElapsedSec** .

For details about implicit units of measure, see [About Units of Measure](http://msdn.microsoft.com/library/b6140312-b8e6-0cf2-9fe0-b14e800216bf%28Office.15%29.aspx).


