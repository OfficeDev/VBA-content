---
title: DoCmd.SetMenuItem Method (Access)
keywords: vbaac10.chm4181
f1_keywords:
- vbaac10.chm4181
ms.prod: access
api_name:
- Access.DoCmd.SetMenuItem
ms.assetid: 690263c1-5e0f-54cd-1032-b2f718d82075
ms.date: 06/08/2017
---


# DoCmd.SetMenuItem Method (Access)

The  **SetMenuItem** method carries out the SetMenuItem action in Visual Basic.


## Syntax

 _expression_. **SetMenuItem**( ** _MenuIndex_**, ** _CommandIndex_**, ** _SubcommandIndex_**, ** _Flag_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



| <strong>Name</strong>                                                                                                                                                                                                                                                                                                                                                                                                                                               | <strong>Required/Optional</strong> | <strong>Data Type</strong> | <strong>Description</strong>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            |
|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|:-----------------------------------|:---------------------------|:----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>MenuIndex</em>                                                                                                                                                                                                                                                                                                                                                                                                                                                  | Required                           | <strong>Variant</strong>   | An integer, counting from 0, that's the valid index of a menu on the custom menu bar or global menu bar for the active window, as defined in the menu bar macro for the custom menu bar or global menu bar. If you select a menu with this argument and leave the commandindex and subcommandindex arguments blank (or set them to ?1), you can enable or disable the menu name itself. You can't, however, check or uncheck a menu name (Microsoft Access ignores the  <strong>acMenuCheck</strong> and <strong>acMenuUncheck</strong> settings for the flag argument for menu names). |
| <em>CommandIndex</em>                                                                                                                                                                                                                                                                                                                                                                                                                                               | Optional                           | <strong>Variant</strong>   | An integer, counting from 0, that's the valid index of a command on the menu selected by the menuindex argument, as defined in the macro group that defines the selected menu for the custom menu bar or global menu bar for the active window.                                                                                                                                                                                                                                                                                                                                         |
| <em>SubcommandIndex</em>                                                                                                                                                                                                                                                                                                                                                                                                                                            | Optional                           | <strong>Variant</strong>   | An integer, counting from 0, that's the valid index of a subcommand in the submenu selected by the commandindex argument, as defined in the macro group that defines the selected submenu for the custom menu bar or global menu bar for the active window.                                                                                                                                                                                                                                                                                                                             |
| <em>Flag</em>                                                                                                                                                                                                                                                                                                                                                                                                                                                       | Optional                           | <strong>Variant</strong>   | The state you want to set the command or subcommand to. Can be one of the following constants.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          |
| <ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p><b>acMenuCheck</b></p></li><li><p><b>acMenuGray</b></p></li><li><p><b>acMenuUncheck</b></p></li><li><p><b>acMenuUngray</b>  (default)</p></li></ul> |                                    |                            |                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |

## Remarks

You can use the  **SetMenuItem** method to set the state of menu items (enabled or disabled, checked or unchecked) on the custom menu bar or global menu bar for the active window.


 **Note**   The **SetMenuItem** method works only with custom menu bars and global menu bars created by using menu bar macros. The **SetMenuItem** method is included in this version of Microsoft Access only for compatibility with versions prior to Microsoft Access 97. It doesn't work with the new command bars functionality.


## See also


#### Concepts


[DoCmd Object](docmd-object-access.md)

