---
title: Form.MenuBar Property (Access)
keywords: vbaac10.chm13385
f1_keywords:
- vbaac10.chm13385
ms.prod: access
api_name:
- Access.Form.MenuBar
ms.assetid: b9e6b6f6-5e60-271d-67c4-6697cb294671
ms.date: 06/08/2017
---


# Form.MenuBar Property (Access)

Specifies a custom menu to display for a form. Read/write  **String**.


## Syntax

 _expression_. **MenuBar**

 _expression_ A variable that represents a **Form** object.


## Remarks

When opening a form in Access that is part of a database that was created in an earlier version of Microsoft Access, the specified menu bar will be displayed differently depending on the curent settings of the  **AllowFullMenus** and **AllowBuiltInToolbars** properties. If the **AllowFullMenus** and **AllowBuiltInToolbars** properties are set to False, the specified menu bar will replace the ribbon as the default set of commands available to the user. If the **AllowFullMenus** or **AllowBuiltInToolbars** property is set to **True**, then the specified menu bar is displayed on the ribbon **Add-Ins** tab.


## See also


#### Concepts


[Form Object](form-object-access.md)

