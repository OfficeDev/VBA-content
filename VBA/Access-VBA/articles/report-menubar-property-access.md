---
title: Report.MenuBar Property (Access)
keywords: vbaac10.chm13710
f1_keywords:
- vbaac10.chm13710
ms.prod: access
api_name:
- Access.Report.MenuBar
ms.assetid: 008e1d2e-f467-05a4-d246-eba85fd626ba
ms.date: 06/08/2017
---


# Report.MenuBar Property (Access)

Specifies a custom menu to display for a report. Read/write  **String**.


## Syntax

 _expression_. **MenuBar**

 _expression_ A variable that represents a **Report** object.


## Remarks

When opening a report in Access that's part of a database that was created in an earlier version of Microsoft Access, the specified menu bar will be displayed differently depending on the current settings of the  **AllowFullMenus** and **AllowBuiltInToolbars** properties. If the **AllowFullMenus** and **AllowBuiltInToolbars** properties are set to False, the specified menu bar will replace the ribbon as the default set of commands available to the user. If the **AllowFullMenus** or **AllowBuiltInToolbars** property is set to **True**, then the specified menu bar is displayed on the ribbon **Add-Ins** tab.


## See also


#### Concepts


[Report Object](report-object-access.md)

