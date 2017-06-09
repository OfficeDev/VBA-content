---
title: Error Function
keywords: vblr6.chm1008914
f1_keywords:
- vblr6.chm1008914
ms.prod: office
ms.assetid: f0e51ff6-34f4-43be-ffcb-d935fa0513c7
ms.date: 06/08/2017
---


# Error Function



Returns the error message that corresponds to a given [error number](vbe-glossary.md).
 **Syntax**
 **Error** [ **(**_errornumber_**)** ]
The optional  _errornumber_[argument](vbe-glossary.md) can be any valid error number. If _errornumber_ is a valid error number, but is not defined, **Error** returns the string "Application-defined or object-defined error." If _errornumber_ is not valid, an error occurs. If _errornumber_ is omitted, the message corresponding to the most recent[run-time error](vbe-glossary.md) is returned. If no run-time error has occurred, or _errornumber_ is 0, **Error** returns a zero-length string ("").
 **Remarks**
Examine the [property](vbe-glossary.md) settings of the **Err** object to identify the most recent run-time error. The return value of the **Error** function corresponds to the **Description** property of the **Err** object.

## Example

This example uses the  **Error** function to print error messages that correspond to the specified error numbers.


```vb
Private Sub PrintError()
    Dim ErrorNumber As Long, count As Long
    count = 1: ErrorNumber = 1
    On Error GoTo EOSb
    Do While count < 100
        Do While Error(ErrorNumber) = "Application-defined or object-defined error": ErrorNumber = ErrorNumber + 1: Loop
        Debug.Print count & "-Error(" & ErrorNumber & "): " & Error(ErrorNumber)
        ErrorNumber = ErrorNumber + 1
        count = count + 1
    Loop
EOSb: Debug.Print ErrorNumber
End Sub


```


