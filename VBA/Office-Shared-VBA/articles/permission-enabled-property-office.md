---
title: Permission.Enabled Property (Office)
keywords: vbaof11.chm261008
f1_keywords:
- vbaof11.chm261008
ms.prod: office
api_name:
- Office.Permission.Enabled
ms.assetid: e77fab6f-0191-3ba4-d418-dc25dc79422d
ms.date: 06/08/2017
---


# Permission.Enabled Property (Office)

Gets or sets a  **Boolean** value that indicates whether permissions are enabled on the active document. Read/write.


## Syntax

 _expression_. **Enabled**

 _expression_ Required. A variable that represents a **[Permission](permission-object-office.md)** object.


## Remarks

Use the  **Enabled** property to determine whether permissions are restricted on the active document, and to enable or disable permissions. Set Enabled to **False** to disable permissions and to remove all users, other than the document author, and their permissions.

When permissions are disabled, the  **Count** property of the **Permission** object returns 0 (zero); however, when permissions are re-enabled, the permissions of the document author remain intact.


## See also


#### Concepts


[Permission Object](permission-object-office.md)
#### Other resources


[Permission Object Members](permission-members-office.md)

