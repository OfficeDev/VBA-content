---
title: InvisibleApp.DefaultAngleUnits Property (Visio)
keywords: vis_sdr.chm17551050
f1_keywords:
- vis_sdr.chm17551050
ms.prod: visio
api_name:
- Visio.InvisibleApp.DefaultAngleUnits
ms.assetid: 5c7f775c-9e2b-10e0-cbc0-2ac0b922ed1a
ms.date: 06/08/2017
---


# InvisibleApp.DefaultAngleUnits Property (Visio)

Determines the default unit of measure for quantities that represent angles. Read/write.


## Syntax

 _expression_ . **DefaultAngleUnits**

 _expression_ A variable that represents an **InvisibleApp** object.


### Return Value

Variant


## Remarks

The  **DefaultAngleUnits** property corresponds to the value shown in the **Angle** box under **Display** on the **Advanced** tab of the **Visio Options** dialog box (click the **File** tab, and then click **Options**).

The return value contains one of the values of  **[VisUnitCodes](visunitcodes-enumeration-visio.md)** , which are declared in the Microsoft Visio type library.

You can specify the value of the  **DefaultAngleUnits** property as an integer (a member of **[VisUnitCodes](visunitcodes-enumeration-visio.md)** ) or a string value such as "degrees". If the string is invalid or the unit code is inappropriate (non-angular), an error is generated.

For a complete list of valid unit strings along with corresponding Automation constants (integer values), see [About Units of Measure](http://msdn.microsoft.com/library/b6140312-b8e6-0cf2-9fe0-b14e800216bf%28Office.15%29.aspx).

Cell formulas that contain a specific unit of measure are displayed in those units regardless of the default angle units setting. Many cell formulas, however, use implicit unit syntax and are displayed in default units.

A program can create a cell whose formula is displayed in default units by setting the cell's  **Formula** property to a string in implicit unit syntax. For example, if the formula for the angle of a shape is "=90[deg,A]" , the result is displayed as "90 deg." if the **DefaultAngleUnits** property is **visDegrees** , and "1.5708 rad." if the **DefaultAngleUnits** property is **visRadians** .

Alternatively, a program can use the following statement to set the cell's result to default angle units:




```
vsoCell.Result(visAngleUnits) = 90
```

In this case, the result is 90 degrees if the  **DefaultAngleUnits** property is **visDegrees** , and 90 radians if the **DefaultAngleUnits** property is **visRadians** .

For details about implicit units of measure, see [About Units of Measure](http://msdn.microsoft.com/library/b6140312-b8e6-0cf2-9fe0-b14e800216bf%28Office.15%29.aspx).


