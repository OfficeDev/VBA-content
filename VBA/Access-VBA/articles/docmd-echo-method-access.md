---
title: DoCmd.Echo Method (Access)
keywords: vbaac10.chm4149
f1_keywords:
- vbaac10.chm4149
ms.prod: access
api_name:
- Access.DoCmd.Echo
ms.assetid: 519b4fe7-ff48-7ab3-3117-43da2278aa66
ms.date: 06/08/2017
---


# DoCmd.Echo Method (Access)

Carries out the Echo action in Visual Basic.


## Syntax

 _expression_. **Echo**( ** _EchoOn_**, ** _StatusBarText_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _EchoOn_|Required|**Variant**|Use  **True** to turn echo on and **False** to turn it off.|
| _StatusBarText_|Optional|**Variant**|A string expression indicating the text that appears in the status bar.|

## Remarks

If you leave the  _StatusBarText_ argument blank, do not use a comma following the _echoon_ argument.

If you turn echo off in Visual Basic, you must turn it back on or it will remain off, even if the user presses CTRL+BREAK or if Visual Basic encounters a breakpoint. You may want to create a macro that turns echo on and then assign that macro to a key combination or a custom menu command. You could then use the key combination or menu command to turn echo on if it has been turned off in Visual Basic.

The  **Echo** method of the **DoCmd** object was added to provide backward compatibility for running the Echo action in Visual Basic code in Microsoft Access for Windows 95. It is recommended that you use the existing **Echo** method of the **Application** object instead.


 **Note**  The  **Echo** method does not affect the visibility of the ribbon or the availability of ribbon commands.


## See also


#### Concepts


[DoCmd Object](docmd-object-access.md)

