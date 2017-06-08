---
title: Form.DataEntry Property (Access)
keywords: vbaac10.chm13359,vbaac10.chm4316
f1_keywords:
- vbaac10.chm13359,vbaac10.chm4316
ms.prod: access
api_name:
- Access.Form.DataEntry
ms.assetid: 0a970904-10f9-d0c3-24d1-0b988725bb38
ms.date: 06/08/2017
---


# Form.DataEntry Property (Access)

You can use the  **DataEntry** property to specify whether a bound form opens to allow data entry only. The **Data Entry** property doesn't determine whether records can be added; it only determines whether existing records are displayed. Read/write **Boolean**.


## Syntax

 _expression_. **DataEntry**

 _expression_ A variable that represents a **Form** object.


## Remarks

This property can be set in any view.

The  **DataEntry** property has an effect only when the **AllowAdditions** property is set to Yes.

Setting the  **DataEntry** property to Yes by using Visual Basic has the same effect as clicking **Data Entry** on the **Records** menu. Setting it to No by using Visual Basic is equivalent to clicking **Remove Filter/Sort** on the **Records** menu.


## See also


#### Concepts


[Form Object](form-object-access.md)

