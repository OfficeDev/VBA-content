---
title: DoCmd.AddMenu Method (Access)
keywords: vbaac10.chm4141
f1_keywords:
- vbaac10.chm4141
ms.prod: access
api_name:
- Access.DoCmd.AddMenu
ms.assetid: d2db2143-fd15-56b3-ee99-b895bc6b21f8
ms.date: 06/08/2017
---


# DoCmd.AddMenu Method (Access)

The  **AddMenu** method carries out the AddMenu action in Visual Basic.


## Syntax

 _expression_. **AddMenu**( ** _MenuName_**, ** _MenuMacroName_**, ** _StatusBarText_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _MenuName_|Required|**Variant**|A string expression that's the valid name of a drop-down menu to add to the custom menu bar or global menu bar. To create an access key so that you can use the keyboard to choose the menu, type an ampersand (&;) before the letter you want to be the access key. This letter will be underlined in the menu name on the menu bar.|
| _MenuMacroName_|Required|**Variant**|A string expression that's the valid name of the macro group that contains the macros for the menu's commands. This is a required argument.|
| _StatusBarText_|Required|**Variant**|A string expression that's the text to display in the status bar when the menu is selected.|

## Remarks

You can use the AddMenu action to create:


- A custom menu bar for a form or report. The custom menu bar replaces the built-in menu bar for the form or report.
    
- A custom shortcut menu for a form, form control, or report. The custom shortcut menu replaces the built-in shortcut menu for the form, form control, or report.
    
- A global menu bar. The global menu bar replaces the built-in menu bar for all Microsoft Access windows, except where you've added a custom menu bar for a form or report.
    
- A global shortcut menu. The global shortcut menu replaces the built-in shortcut menu for fields in table and query datasheets, forms in Form view, Datasheet view, and Print Preview, and reports in Print Preview, except where you've added a custom shortcut menu for a form, form control, or report.
    
You must include the  _menuname_ and _menumacroname_ arguments in the **AddMenu** method for custom menu bars and global menu bars. The _menuname_ argument is not required and will be ignored for custom shortcut menus and global shortcut menus.

The  _statusbartext_ argument is optional, this argument is ignored for custom shortcut menus and global shortcut menus.


## See also


#### Concepts


[DoCmd Object](docmd-object-access.md)

