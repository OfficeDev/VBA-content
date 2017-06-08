---
title: Application.MenuBar Property (Access)
keywords: vbaac10.chm12500
f1_keywords:
- vbaac10.chm12500
ms.prod: access
api_name:
- Access.Application.MenuBar
ms.assetid: dc0f6f9c-4627-96a1-83fa-b58ce1eb7236
ms.date: 06/08/2017
---


# Application.MenuBar Property (Access)

Specifies a custom menu to display for a Microsoft Access database. Read/write  **String**.


## Syntax

 _expression_. **MenuBar**

 _expression_ A variable that represents an **Application** object.


## Remarks

When opening a database in Access that was created in an earlier version of Microsoft Access, the specified menu bar will be displayed differently depending on the curent settings of the  **AllowFullMenus** and **AllowBuiltInToolbars** properties. If the **AllowFullMenus** and **AllowBuiltInToolbars** properties are set to **False**, the specified menu bar will replace the ribbon as the default set of commands available to the user. If the **AllowFullMenus** or **AllowBuiltInToolbars** property is set to **True**, then the specified menu bar is displayed on the ribbon **Add-Ins** tab.


## See also


#### Concepts


[Application Object](application-object-access.md)

