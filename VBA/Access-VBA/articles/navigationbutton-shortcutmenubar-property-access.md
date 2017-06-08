---
title: NavigationButton.ShortcutMenuBar Property (Access)
keywords: vbaac10.chm10479
f1_keywords:
- vbaac10.chm10479
ms.prod: access
api_name:
- Access.NavigationButton.ShortcutMenuBar
ms.assetid: bfc92fea-48ef-e995-53c4-be0354de1550
ms.date: 06/08/2017
---


# NavigationButton.ShortcutMenuBar Property (Access)

You can use the  **ShortcutMenuBar** property to specify the shortcut menu that will appear when you right-click on the specified object. Read/write **String**.


## Syntax

 _expression_. **ShortcutMenuBar**

 _expression_ A variable that represents a **NavigationButton** object.


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


[NavigationButton Object](navigationbutton-object-access.md)

