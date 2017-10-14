---
title: InvisibleApp.DefaultTextUnits Property (Visio)
keywords: vis_sdr.chm17551035
f1_keywords:
- vis_sdr.chm17551035
ms.prod: visio
api_name:
- Visio.InvisibleApp.DefaultTextUnits
ms.assetid: a9bd8464-b39c-622c-6446-bc652e42766c
ms.date: 06/08/2017
---


# InvisibleApp.DefaultTextUnits Property (Visio)

Determines the default unit of measure for quantities that represent text metrics. Read/write.


## Syntax

 _expression_ . **DefaultTextUnits**

 _expression_ A variable that represents an **InvisibleApp** object.


### Return Value

Variant


## Remarks

The  **DefaultTextUnits** property corresponds to the value shown in the **Text** box under **Display** on the **Advanced** tab of the **Visio Options** dialog box (click the **File** tab, and then click **Options**).

The return value contains one of the values of  **[VisUnitCodes](visunitcodes-enumeration-visio.md)** , which are declared in the Microsoft Visio type library.

You can specify the value of  **DefaultTextUnits** as an integer (a member of **[VisUnitCodes](visunitcodes-enumeration-visio.md)** ) or a string value such as "pt". If the string is invalid or the unit code is inappropriate (non-textual), an error is generated.

For a complete list of valid unit strings along with corresponding Automation constants (integer values), see [About Units of Measure](http://msdn.microsoft.com/library/b6140312-b8e6-0cf2-9fe0-b14e800216bf%28Office.15%29.aspx).

Cell formulas that contain a specific unit of measure are displayed in those units regardless of the default text units setting. Many cell formulas, however, use implicit unit syntax and are displayed in default units.

A program can create a cell whose formula is displayed in default units by setting the cell's  **Formula** property to a string in implicit unit syntax. For example, the formula "=8[pt,T]" is displayed as "8 pt" if the **DefaultTextUnits** property is **visPoints** and "0.6272" if the **DefaultTextUnits** property is **visCiceros** .

Alternatively, a program can use the following statement to set the cell's result to default text units: 




```
vsoCell.Result(visTextUnits) = 12
```

In this case, the text is 12 points if the  **DefaultTextUnits** property is **visPoints** and 12 ciceros if the **DefaultTextUnits** property is **visCiceros** .

For details about implicit units of measure, see [About Units of Measure](http://msdn.microsoft.com/library/b6140312-b8e6-0cf2-9fe0-b14e800216bf%28Office.15%29.aspx).


