---
title: Form.ShortcutMenu Property (Access)
keywords: vbaac10.chm13387,vbaac10.chm4502
f1_keywords:
- vbaac10.chm13387,vbaac10.chm4502
ms.prod: access
api_name:
- Access.Form.ShortcutMenu
ms.assetid: ec652f43-4dc8-4970-19ad-d117c3193528
ms.date: 06/08/2017
---


# Form.ShortcutMenu Property (Access)

You can use the  **ShortcutMenu** property to specify whether a shortcut menu is displayed when you right-click an object on a form. For example, you might want to disable a shortcut menu to prevent the user from changing the form's underlying record source by using one of the filtering commands on the form's shortcut menu. Read/write **Boolean**.


## Syntax

 _expression_. **ShortcutMenu**

 _expression_ A variable that represents a **Form** object.


## Remarks

The default value is  **True**.

This property controls the displaying of the shortcut menus for a form and for any of its controls. If the  **ShortcutMenu** property is set to **False**, shortcut menus aren't displayed when you right-click a form or any of its controls.

If you're developing a wizard, you might want to hide shortcut menus on your wizard forms to prevent the user from viewing or using them. This is especially true for forms that display choices. For example, the  **ShortcutMenu** property for the Startup form in the Northwind sample database is set to No. This prevents users from displaying shortcut menus for the form or controls on the form.


## Example

The following example disables the shortcut menus for the Invoice form and its controls:


```vb
Forms!Invoice.ShortcutMenu = False
```


## See also


#### Concepts


[Form Object](form-object-access.md)

