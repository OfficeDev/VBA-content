---
title: Application.Quit Method (Access)
keywords: vbaac10.chm12507
f1_keywords:
- vbaac10.chm12507
ms.prod: access
api_name:
- Access.Application.Quit
ms.assetid: 075ad885-f25d-ea2d-bf74-8ec915265c63
ms.date: 06/08/2017
---


# Application.Quit Method (Access)

The [Quit](application-quit-method-access.md) method quits Microsoft Access. You can select one of several options for saving a database object before quitting.


## Syntax

 _expression_. **Quit**( ** _Option_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Options_|Optional|**AcQuitOption**|An [AcQuitOption](acquitoption-enumeration-access.md) constant that indicates the action to take when quitting Access. The default value is **acQuitSaveAll**.|

## See also


#### Concepts


[Application Object](application-object-access.md)

