---
title: Worksheet.Delete Method (Excel)
keywords: vbaxl10.chm174075
f1_keywords:
- vbaxl10.chm174075
ms.prod: excel
api_name:
- Excel.Worksheet.Delete
ms.assetid: a51e1673-e09d-824f-1acc-dda18c120204
ms.date: 06/08/2017
---


# Worksheet.Delete Method (Excel)

Deletes the object.


## Syntax

 _expression_ . **Delete**

 _expression_ A variable that represents a **Worksheet** object.


### Return Value

Boolean


## Remarks

When you delete a  **[Worksheet](worksheet-object-excel.md)** , this method displays a dialog box that prompts the user to confirm the deletion. This dialog box is displayed by default. When called on the **Worksheet** object, the **Delete** method returns a **Boolean** value that is **False** if the user clicked **Cancel** on the dialog box or **True** if the user clicked **Delete**.


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

