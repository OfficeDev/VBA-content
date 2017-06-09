---
title: HelpContext Property (Visual Basic for Applications)
keywords: vblr6.chm1014190
f1_keywords:
- vblr6.chm1014190
ms.prod: office
ms.assetid: 5cfd1f6c-1d91-623c-dbb0-3431d5837881
ms.date: 06/08/2017
---


# HelpContext Property (Visual Basic for Applications)



Returns or sets a [string expression](vbe-glossary.md) containing the context ID for a topic in a Help file. Read/write.
 **Remarks**
The  **HelpContext**[property](vbe-glossary.md) is used to automatically display the Help topic specified in the **HelpFile** property. If both **HelpFile** and **HelpContext** are empty, the value of **Number** is checked. If **Number** corresponds to a Visual Basic[run-time error](vbe-glossary.md) value, then the Visual Basic Help context ID for the error is used. If the **Number** value doesn't correspond to a Visual Basic error, the contents screen for the Visual Basic Help file is displayed.

 **Note**  You should write routines in your application to handle typical errors. When programming with an object, you can use the object's Help file to improve the quality of your error handling, or to display a meaningful message to your user if the error isn't recoverable.


## Example

This example uses the  **HelpContext** property of the **Err** object to show the Visual Basic Help topic for the `Overflow` error.


```vb
Dim Msg
Err.Clear
On Error Resume Next
Err.Raise 6 ' Generate "Overflow" error.
If Err.Number <> 0 Then
    Msg = "Press F1 or HELP to see " &; Err.HelpFile &; " topic for" &; _
    " the following HelpContext: " &; Err. HelpContext
    MsgBox Msg, , "Error: " &; Err.Description, Err.HelpFile, _
Err.HelpContext
End If
```


