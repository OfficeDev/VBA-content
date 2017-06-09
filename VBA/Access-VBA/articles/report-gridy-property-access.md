---
title: Report.GridY Property (Access)
keywords: vbaac10.chm13714
f1_keywords:
- vbaac10.chm13714
ms.prod: access
api_name:
- Access.Report.GridY
ms.assetid: e4a13708-fa05-8ac4-af5f-0f78ee15e623
ms.date: 06/08/2017
---


# Report.GridY Property (Access)

You can use the  **GridY** property (along with the **GridX** property) to specify the horizontal and vertical divisions of the alignment grid in report Design view. Read/write **Integer**.


## Syntax

 _expression_. **GridY**

 _expression_ A variable that represents a **Report** object.


## Remarks

Enter an integer between 1 and 64 representing the number of subdivisions per unit of measurement. If the  **Measurement system** box is set to U.S. on the **Numbers** tab of the **Regional Options** dialog box of Windows Control Panel, the default setting is 24 for the **GridX** property (horizontal) and 24 for the **GridY** property (vertical).

In Visual Basic, you set this property by using a numeric expression.

The  **GridX** and **GridY** properties provide control over the placement and alignment of objects on a form or report. You can adjust the grid for greater or lesser precision. To see the grid, click **Grid** on the **View** menu. If the setting for either the **GridX** or **GridY** properties is greater than 24, the grid points disappear from view (although the grid lines are still displayed).


## See also


#### Concepts


[Report Object](report-object-access.md)

