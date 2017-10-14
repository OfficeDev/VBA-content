---
title: Application.ShortcutMenuBar Property (Access)
keywords: vbaac10.chm12512
f1_keywords:
- vbaac10.chm12512
ms.prod: access
api_name:
- Access.Application.ShortcutMenuBar
ms.assetid: 6785320b-b50f-dcaa-3eae-13d378573613
ms.date: 06/08/2017
---


# Application.ShortcutMenuBar Property (Access)

You can use the  **ShortcutMenuBar** property to specify the shortcut menu that will appear when you right-click on the specified object. Read/write **String**.


## Syntax

 _expression_. **ShortcutMenuBar**

 _expression_ A variable that represents an **Application** object.


## Remarks

When used with the  **[Application](application-object-access.md)** object, the **ShortcutMenuBar** property enables you to display a custom shortcut menu as a global shortcut menu. However, if you've set the **ShortcutMenuBar** property for a form, form control, or report in the database, the custom shortcut menu of that object will be displayed in place of the database's global shortcut menu. You can display a different custom shortcut menu for a specific form, form control, or report by setting its **ShortcutMenuBar** property to a different shortcut menu. When the form, form control, or report has the focus, the custom shortcut menu for that object is displayed when the user clicks the right mouse button; otherwise, the global shortcut menu for the database is displayed.

Shortcut menus aren't available to any object if the  **AllowShortcutMenus** property is set to **False**.


## See also


#### Concepts


[Application Object](application-object-access.md)

