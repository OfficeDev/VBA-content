---
title: Raise Method
keywords: vblr6.chm1014183
f1_keywords:
- vblr6.chm1014183
ms.prod: office
api_name:
- Office.Raise
ms.assetid: 7e3ddb06-db93-ebce-7562-8a15c49261b1
ms.date: 06/08/2017
---


# Raise Method



Generates a [run-time error](vbe-glossary.md).
 **Syntax**
 _object_**.Raise  _number_,** **_source_, _description_, _helpfile_, _helpcontext_**
The  **Raise** method has the following object qualifier and[named arguments](vbe-glossary.md):


|**Argument**|**Description**|
|:-----|:-----|
| _object_|Required. Always the  **Err** object.|
|**_number_**|Required. [Long](vbe-glossary.md) integer that identifies the nature of the error. Visual Basic errors (both Visual Basic-defined and user-defined errors) are in the range 0-65535. The range 0-512 is reserved for system errors; the range 513-65535 is available for user-defined errors. When setting the **Number** property to your own error code in a class module, you add your error code number to the **vbObjectError**[constant](vbe-glossary.md). For example, to generate the [error number](vbe-glossary.md) 513, assign **vbObjectError** + 513 to the **Number** property.|
|**_source_**|Optional. [String expression](vbe-glossary.md) naming the object or application that generated the error. When setting this[property](vbe-glossary.md) for an object, use the form _project.class_. If _source_ is not specified, the programmatic ID of the current Visual Basic[project](vbe-glossary.md) is used.|
|**_description_**|Optional. String expression describing the error. If unspecified, the value in  **Number** is examined. If it can be mapped to a Visual Basic run-time error code, the string that would be returned by the **Error** function is used as **Description**_._ If there is no Visual Basic error corresponding to **Number**, the "Application-defined or object-defined error" message is used.|
|**_helpfile_**|Optional. The fully qualified path to the Help file in which help on this error can be found. If unspecified, Visual Basic uses the fully qualified drive, path, and file name of the Visual Basic Help file.|
|**_helpcontext_**|Optional. The context ID identifying a topic within  **_helpfile_** that provides help for the error. If omitted, the Visual Basic Help file context ID for the error corresponding to the **Number** property is used, if it exists.|
 **Remarks**
All of the [arguments](vbe-glossary.md) are optional except **_number_**. If you use **Raise** without specifying some arguments, and the property settings of the **Err** object contain values that have not been cleared, those values serve as the values for your error.
 **Raise** is used for generating run-time errors and can be used instead of the **Error** statement. **Raise** is useful for generating errors when writing class modules, because the **Err** object gives richer information than is possible if you generate errors with the **Error** statement. For example, with the **Raise** method, the source that generated the error can be specified in the **Source** property, online Help for the error can be referenced, and so on.

## Example

This example uses the  **Err** object's **Raise** method to generate an error within an Automation object written in Visual Basic. It has the programmatic ID `MyProj.MyObject`. On the MacIntosh, the default drive name is "HD" and portions of the pathname are separated by colons instead of backslashes.


```vb
Const MyContextID = 1010407    ' Define a constant for contextID.
Function TestName(CurrentName, NewName)
    If Instr(NewName, "bob") Then    ' Test the validity of NewName.
        ' Raise the exception
        Err.Raise vbObjectError + 513, "MyProj.MyObject", _
        "No ""bob"" allowed in your name", "c:\MyProj\MyHelp.Hlp", _
        MyContextID
    End If
End Function
```


