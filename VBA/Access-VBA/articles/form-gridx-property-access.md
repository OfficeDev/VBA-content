---
title: Form.GridX Property (Access)
keywords: vbaac10.chm13389
f1_keywords:
- vbaac10.chm13389
ms.prod: access
api_name:
- Access.Form.GridX
ms.assetid: ebc6a4d9-2f73-cf55-504f-a83aff1fecd4
ms.date: 06/08/2017
---


# Form.GridX Property (Access)

You can use the  **GridX** property (along with the **GridY** property) to specify the horizontal and vertical divisions of the alignment grid in form Design view. Read/write **Integer**.


## Syntax

 _expression_. **GridX**

 _expression_ A variable that represents a **Form** object.


## Remarks

Enter an integer between 1 and 64 representing the number of subdivisions per unit of measurement. If the  **Measurement system** box is set to U.S. on the **Numbers** tab of the **Regional Options** dialog box of Windows Control Panel, the default setting is 24 for the **GridX** property (horizontal) and 24 for the **GridY** property (vertical).

In Visual Basic, you set this property by using a numeric expression.

The  **GridX** and **GridY** properties provide control over the placement and alignment of objects on a form or report. You can adjust the grid for greater or lesser precision. To see the grid, click **Grid** on the **View** menu. If the setting for either the **GridX** or **GridY** properties is greater than 24, the grid points disappear from view (although the grid lines are still displayed).


## See also


#### Concepts


[Form Object](form-object-access.md)

