---
title: Number Property (Visual Basic for Applications)
keywords: vblr6.chm1014186
f1_keywords:
- vblr6.chm1014186
ms.prod: office
ms.assetid: e6a671d8-0fd3-9d92-efd2-bcf3d0ff0758
ms.date: 06/08/2017
---


# Number Property (Visual Basic for Applications)



Returns or sets a numeric value specifying an error.  **Number** is the **Err** object's default property. Read/write.
 **Remarks**
When returning a user-defined error from an object, set  **Err.Number** by adding the number you selected as an error code to the **vbObjectError**[constant](vbe-glossary.md). For example, you use the following code to return the number 1051 as an error code:



```
Err.Raise Number := vbObjectError + 1051, Source:= "SomeClass"


```


## Example

The first example illustrates a typical use of the  **Number** property in an error-handling routine. The second example examines the **Number** property of the **Err** object to determine whether an error returned by an Automation object was defined by the object, or whether it was mapped to an error defined by Visual Basic. Note that the constant **vbObjectError** is a very large negative number that an object adds to its own error code to indicate that the error is defined by the server. Therefore, subtracting it from **Err.Number** strips it out of the result. If the error is object-defined, the base number is left in `MyError`, which is displayed in a message box along with the original source of the error. If  **Err.Number** represents a Visual Basic error, then the Visual Basic error number is displayed in the message box.


```vb
' Typical use of Number property
Sub test()
    On Error GoTo out
    
    Dim x, y
    x = 1 / y    ' Create division by zero error
    Exit Sub
    out:
    MsgBox Err. Number
    MsgBox Err.Description
    ' Check for division by zero error
    If Err.Number = 11 Then
        y = y + 1
    End If
    Resume
End Sub



' Using Number property with an error from an 
' Automation object
Dim MyError, Msg
' First, strip off the constant added by the object to indicate one
' of its own errors.
MyError = Err. Number - vbObjectError
' If you subtract the vbObjectError constant, and the number is still 
' in the range 0-65,535, it is an object-defined error code.
If MyError > 0 And MyError < 65535 Then
    Msg = "The object you accessed assigned this number to the error: " _
             &; MyError &; ". The originator of the error was: " _
            &; Err.Source &; ". Press F1 to see originator's Help topic."
' Otherwise it is a Visual Basic error number.
Else
    Msg = "This error (# " &; Err. Number &; ") is a Visual Basic error" &; _
            " number. Press Help button or F1 for the Visual Basic Help" _
            &; " topic for this error."
End If
    MsgBox Msg, , "Object Error", Err.HelpFile, Err.HelpContext

```


