---
title: Report.Toolbar Property (Access)
keywords: vbaac10.chm13711
f1_keywords:
- vbaac10.chm13711
ms.prod: access
api_name:
- Access.Report.Toolbar
ms.assetid: e897d294-2d8d-aca7-9aed-4bd2ebd23552
ms.date: 06/08/2017
---


# Report.Toolbar Property (Access)

Specifies a custom toolbar to display for a report. Read/write  **String**.


## Syntax

 _expression_. **Toolbar**

 _expression_ A variable that represents a **Report** object.


## Remarks

When opening a report in Access that's part of a database that was created in an earlier version of Microsoft Access, the specified toolbar will be displayed differently depending on the current settings of the  **AllowFullMenus** and **AllowBuiltInToolbars** properties. If the **AllowFullMenus** and **AllowBuiltInToolbars** properties are set to False, the specified toolbar will replace the ribbon as the default set of commands available to the user. If the **AllowFullMenus** or **AllowBuiltInToolbars** property is set to **True**, then the specified toolbar is displayed on the ribbon **Add-Ins** tab.


## See also


#### Concepts


[Report Object](report-object-access.md)

