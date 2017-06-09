---
title: Presentation.Password Property (PowerPoint)
keywords: vbapp10.chm583080
f1_keywords:
- vbapp10.chm583080
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.Password
ms.assetid: 977876b7-b40f-de45-c259-e91744915085
ms.date: 06/08/2017
---


# Presentation.Password Property (PowerPoint)

Returns or sets the password that must be supplied to open the specified presentation. Read/write.


## Syntax

 _expression_. **Password**

 _expression_ A variable that represents a **Presentation** object.


### Return Value

String


## Example

This example opens Earnings.ppt, sets a password for it, and then closes the presentation.


```vb
Sub SetPassword()

    With Presentations.Open(FileName:="C:\My Documents\Earnings.ppt")

        .Password = complexstrPWD 'global variable

        .Save

        .Close

    End With

End Sub
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

