---
title: Attachment.ShortcutMenuBar Property (Access)
keywords: vbaac10.chm13971
f1_keywords:
- vbaac10.chm13971
ms.prod: access
api_name:
- Access.Attachment.ShortcutMenuBar
ms.assetid: be4ce61e-c4a9-9e3b-e2f4-187b77451f67
ms.date: 06/08/2017
---


# Attachment.ShortcutMenuBar Property (Access)

You can use the  **ShortcutMenuBar** property to specify the shortcut menu that will appear when you right-click the specified object. Read/write **String**.


## Syntax

 _expression_. **ShortcutMenuBar**

 _expression_ A variable that represents an **Attachment** object.


## Remarks


 **Note**  The  **ShortcutMenuBar** property applies only to controls on a form, not controls on a report.

You can also use the  **ShortcutMenuBar** property to specify the menu bar macro that will be used to display a shortcut menu for a datasheet, form, form control, or report.

To display the built-in shortcut menu for a database, form, form control, or report by using a macro or Visual Basic, set the property to a zero-length string (" ").

When used with the  **[Application](application-object-access.md)** object, the **ShortcutMenuBar** property enables you to display a custom shortcut menu as a global shortcut menu. However, if you've set the **ShortcutMenuBar** property for a form, form control, or report in the database, the custom shortcut menu of that object will be displayed in place of the database's global shortcut menu. You can display a different custom shortcut menu for a specific form, form control, or report by setting its **ShortcutMenuBar** property to a different shortcut menu. When the form, form control, or report has the focus, the custom shortcut menu for that object is displayed when the user clicks the right mouse button; otherwise, the global shortcut menu for the database is displayed.

Shortcut menus aren't available to any object if the  **AllowShortcutMenus** property is set to **False**.


## See also


#### Concepts


[Attachment Object](attachment-object-access.md)

