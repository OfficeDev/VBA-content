---
title: HelpFile Property
keywords: vblr6.chm1014189
f1_keywords:
- vblr6.chm1014189
ms.prod: office
api_name:
- Office.HelpFile
ms.assetid: ed2b9abc-679c-91d9-4d39-987863430880
ms.date: 06/08/2017
---


# HelpFile Property



Returns or sets a [string expression](vbe-glossary.md) the fully qualified path to a Help file. Read/write.
 **Remarks**
If a Help file is specified in  **HelpFile**, it is automatically called when the user presses the **Help** button (or the F1 KEY in Windows or the HELP key on the Macintosh) in the error message dialog box. If the **HelpContext** property contains a valid context ID for the specified file, that topic is automatically displayed. If no **HelpFile** is specified, the Visual Basic Help file is displayed.

 **Note**  You should write routines in your application to handle typical errors. When programming with an object, you can use the object's Help file to improve the quality of your error handling, or to display a meaningful message to your user if the error isn't recoverable.


## Example

This example uses the  **HelpFile** property of the **Err** object to start the Help system. By default, the **HelpFile** property contains the name of the Visual Basic Help file.


```vb
Dim Msg
Err.Clear
On Error Resume Next    ' Suppress errors for demonstration purposes.
Err.Raise 6    ' Generate "Overflow" error.
Msg = "Press F1 or HELP to see " &; Err. HelpFile &; _
" topic for this error"
MsgBox Msg, , "Error: " &; Err.Description,Err. HelpFile
```


