---
title: Application.ActivePrinter Property (PowerPoint)
keywords: vbapp10.chm502017
f1_keywords:
- vbapp10.chm502017
ms.prod: powerpoint
api_name:
- PowerPoint.Application.ActivePrinter
ms.assetid: 48ba3853-6a8f-d523-807a-8324e59adbb7
ms.date: 06/08/2017
---


# Application.ActivePrinter Property (PowerPoint)

Returns the name of the active printer. Read-only.


## Syntax

 _expression_. **ActivePrinter**

 _expression_ A variable that represents an **Application** object.


### Return Value

String


## Remarks

This example displays the name of the active printer.


## Example

This example displays the name of the active printer.


```vb
MsgBox "The name of the active printer is " &; Application.ActivePrinter
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

