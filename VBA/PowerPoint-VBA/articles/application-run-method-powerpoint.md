---
title: Application.Run Method (PowerPoint)
keywords: vbapp10.chm502023
f1_keywords:
- vbapp10.chm502023
ms.prod: powerpoint
api_name:
- PowerPoint.Application.Run
ms.assetid: 21b8a0c4-10c8-d8c3-9214-adffad35f7d4
ms.date: 06/08/2017
---


# Application.Run Method (PowerPoint)

Runs a Visual Basic procedure.


## Syntax

 _expression_. **Run**( **_MacroName_**, **_safeArrayOfParams_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _MacroName_|Required|**String**|The name of the procedure to be run. The string can contain the following: a loaded presentation or add-in file name followed by an exclamation point (!), a valid module name followed by a period (.), and the procedure name. For example, the following is a valid MacroName value: "MyPres.ppt!Module1.Test."|
| _safeArrayOfParams()_|Required|**Variant**|The argument to be passed to the procedure. You cannot specify an object for this argument, and you cannot use named arguments with this method. Arguments must be passed by position.|

### Return Value

Variant


## Example

In this example, the Main procedure defines an array and then runs the macro TestPass, passing the array as an argument.


```vb
Sub Main()

    Dim x(1 To 2)

    x(1) = "hi"

    x(2) = 7

    Application.Run "TestPass", x

End Sub



Sub TestPass(x)

    MsgBox x(1)

    MsgBox x(2)

End Sub
```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

