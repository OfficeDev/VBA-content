---
title: Error Statement
keywords: vblr6.chm1008913
f1_keywords:
- vblr6.chm1008913
ms.prod: office
ms.assetid: b657920d-b28c-0c6b-8020-9d37e9f10f6c
ms.date: 06/08/2017
---


# Error Statement

Simulates the occurrence of an error.

 **Syntax**

 **Error**_errornumber_

The required  _errornumber_ can be any valid[error number](vbe-glossary.md).
 **Remarks**
The  **Error** statement is supported for backward compatibility. In new code, especially when creating objects, use the **Err** object's **Raise** method to generate[run-time errors](vbe-glossary.md).
If  _errornumber_ is defined, the **Error** statement calls the error handler after the[properties](vbe-glossary.md) of **Err** object are assigned the following default values:


| <strong>Property</strong>     | <strong>Value</strong>                                                                                                                                                                                                                                                        |
|:------------------------------|:------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <strong>Number</strong>       | Value specified as [argument](vbe-glossary.md) to <strong>Error</strong> statement. Can be any valid error number.                                                                                                                                                            |
| <strong>Source</strong>       | Name of the current Visual Basic [project](vbe-glossary.md).                                                                                                                                                                                                                  |
| <strong>Description</strong>  | [String expression](vbe-glossary.md) corresponding to the return value of the <strong>Error</strong> function for the specified <strong>Number</strong>, if this string exists. If the string doesn't exist, <strong>Description</strong> contains a zero-length string (""). |
| <strong>HelpFile</strong>     | The fully qualified drive, path, and file name of the appropriate Visual Basic Help file.                                                                                                                                                                                     |
| <strong>HelpContext</strong>  | The appropriate Visual Basic Help file context ID for the error corresponding to the  <strong>Number</strong> property.                                                                                                                                                       |
| <strong>LastDLLError</strong> | Zero.                                                                                                                                                                                                                                                                         |

If no error handler exists or if none is enabled, an error message is created and displayed from the  **Err** object properties.

 **Note**  Not all Visual Basic [host applications](vbe-glossary.md) can create objects. See your host application's documentation to determine whether it can create[classes](vbe-glossary.md) and objects.


## Example

This example uses the  **Error** statement to simulate error number 11.


```vb
On Error Resume Next ' Defer error handling. 
Error 11 ' Simulate the "Division by zero" error. 
```


