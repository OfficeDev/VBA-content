---
title: Command Function
keywords: vblr6.chm1013109
f1_keywords:
- vblr6.chm1013109
ms.prod: office
ms.assetid: 2eaefb12-2e7f-ab4a-9cd8-fc0739e33bf5
ms.date: 06/08/2017
---


# Command Function



Returns the [argument](vbe-glossary.md) portion of the[command line](vbe-glossary.md) used to launch Microsoft Visual Basic or an executable program developed with Visual Basic. The Visual Basic **Command** function is not available in Microsoft Office applications.
 **Syntax**
 **Command**
 **Remarks**
When Visual Basic is launched from the command line, any portion of the command line that follows  `/cmd` is passed to the program as the command-line argument. In the following example, is passed to the program as the command-line argument. In the following command line example, `cmdlineargs` represents the argument information returned by the **Command** function.



```text
VB /cmd cmdlineargs
```

For applications developed with Visual Basic and compiled to an .exe file,  **Command** returns any arguments that appear after the name of the application on the command line. For example:



```text
MyApp cmdlineargs
```

To find how command line arguments can be changed in the user interface of the application you're using, search Help for "command line arguments."

## Example

This example uses the  **Command** function to get the command line arguments in a function that returns them in a **Variant** containing an array. Not available in Microsoft Office.


```vb
Function GetCommandLine(Optional MaxArgs)
    'Declare variables.
    Dim C, CmdLine, CmdLnLen, InArg, I, NumArgs
    'See if MaxArgs was provided.
    If IsMissing(MaxArgs) Then MaxArgs = 10
    'Make array of the correct size.
    ReDim ArgArray(MaxArgs)
    NumArgs = 0: InArg = False
    'Get command line arguments.
    CmdLine = Command()
    CmdLnLen = Len(CmdLine)
    'Go thru command line one character
    'at a time.
    For I = 1 To CmdLnLen
        C = Mid(CmdLine, I, 1)
        'Test for space or tab.
        If (C <> " " And C <> vbTab) Then
            'Neither space nor tab.
            'Test if already in argument.
            If Not InArg Then
            'New argument begins.
            'Test for too many arguments.
                If NumArgs = MaxArgs Then Exit For
                NumArgs = NumArgs + 1
                InArg = True
            End If
            'Concatenate character to current argument.
            ArgArray(NumArgs) = ArgArray(NumArgs) &; C
        Else
            'Found a space or tab.
            'Set InArg flag to False.
            InArg = False
        End If
    Next I
    'Resize array just enough to hold arguments.
    ReDim Preserve ArgArray(NumArgs)
    'Return Array in Function name.
    GetCommandLine = ArgArray()
End Function
```


