---
title: Page.ShortcutMenuBar Property (Access)
keywords: vbaac10.chm12159
f1_keywords:
- vbaac10.chm12159
ms.prod: access
api_name:
- Access.Page.ShortcutMenuBar
ms.assetid: eb9a1075-27c3-4b37-1ebb-7693afd6ba5e
ms.date: 06/08/2017
---


# Page.ShortcutMenuBar Property (Access)

You can use the  **ShortcutMenuBar** property to specify the shortcut menu that will appear when you right-click on the specified object. Read/write **String**.


## Syntax

 _expression_. **ShortcutMenuBar**

 _expression_ A variable that represents a **Page** object.


## Remarks


 **Note**  The  **ShortcutMenuBar** property applies only to controls on a form, not controls on a report.

You can also use the  **ShortcutMenuBar** property to specify the menu bar macro that will be used to display a shortcut menu for a datasheet, form, form control, or report.

To display the built-in shortcut menu for a database, form, form control, or report by using a macro or Visual Basic, set the property to a zero-length string (" ").

When used with the  **[Application](application-object-access.md)** object, the **ShortcutMenuBar** property enables you to display a custom shortcut menu as a global shortcut menu. However, if you've set the **ShortcutMenuBar** property for a form, form control, or report in the database, the custom shortcut menu of that object will be displayed in place of the database's global shortcut menu. You can display a different custom shortcut menu for a specific form, form control, or report by setting its **ShortcutMenuBar** property to a different shortcut menu. When the form, form control, or report has the focus, the custom shortcut menu for that object is displayed when the user clicks the right mouse button; otherwise, the global shortcut menu for the database is displayed.

Shortcut menus aren't available to any object if the  **AllowShortcutMenus** property is set to **False**.


## Example

The following example sets the "Suppliers_Toolbar" as the custom shortcut menu to display when the user clicks the right mouse button on the "Suppliers" form.


```vb
Forms("Suppliers").ShortcutMenuBar = "Suppliers_Toolbar"
```


## See also


#### Concepts


[Page Object](page-object-access.md)

