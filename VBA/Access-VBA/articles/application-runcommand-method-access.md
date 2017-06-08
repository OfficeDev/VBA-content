---
title: Application.RunCommand Method (Access)
keywords: vbaac10.chm12568
f1_keywords:
- vbaac10.chm12568
ms.prod: access
api_name:
- Access.Application.RunCommand
ms.assetid: 2731352f-7f2d-db3a-314c-e8a789755dd5
ms.date: 06/08/2017
---


# Application.RunCommand Method (Access)

The  **RunCommand** method runs a built-in command.


## Syntax

 _expression_. **RunCommand**( ** _Command_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Command_|Required|**AcCommand**|An  **[AcCommand](accommand-enumeration-access.md)** constant that specifies the commend to run.|

## Remarks

Each menu and toolbar command in Microsoft Access has an associated constant that you can use with the  **RunCommand** method to run that command from Visual Basic.

You can't use the  **RunCommand** method to run a command on a custom menu or toolbar. You can only use it with built-in menus and toolbars.

The  **RunCommand** method replaces the **DoMenuItem** method of the **DoCmd** object.


## See also


#### Concepts


[Application Object](application-object-access.md)

