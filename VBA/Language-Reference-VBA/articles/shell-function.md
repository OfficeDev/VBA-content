---
title: Shell Function
keywords: vblr6.chm1009023
f1_keywords:
- vblr6.chm1009023
ms.prod: office
ms.assetid: 033bffb0-540f-2c17-2aed-d25d10bedd8c
ms.date: 06/08/2017
---


# Shell Function



Runs an executable program and returns a  **Variant** ( **Double** ) representing the program's task ID if successful, otherwise it returns zero.
 **Syntax**
 **Shell( _pathname_** [ **, _windowstyle_** ] **)**
The  **Shell** function syntax has these[named arguments](vbe-glossary.md):


| <strong>Part</strong>                 | <strong>Description</strong>                                                                                                                                                                                                                                                                                                                                                                                                                        |
|:--------------------------------------|:----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <strong><em>pathname</em></strong>    | Required;  <strong>Variant</strong> ( <strong>String</strong> ). Name of the program to execute and any required[arguments](vbe-glossary.md) or[command-line](vbe-glossary.md) switches; may include directory or folder and drive. On the Macintosh, you can use the <strong>MacID</strong> function to specify an application's signature instead of its name. The following example uses the signature for Microsoft Word: `Shell MacID("MSWD")` |
| <strong><em>windowstyle</em></strong> | Optional.  <strong>Variant</strong> ( <strong>Integer</strong> ) corresponding to the style of the window in which the program is to be run. If <strong><em>windowstyle</em></strong> is omitted, the program is started minimized with focus. On the Macintosh (System 7.0 or later), <strong><em>windowstyle</em></strong> only determines whether or not the application gets the focus when it is run.                                          |

The  **_windowstyle_** named argument has these values:


| <strong>Constant</strong>           | <strong>Value</strong> | <strong>Description</strong>                                                                                                               |
|:------------------------------------|:-----------------------|:-------------------------------------------------------------------------------------------------------------------------------------------|
| <strong>vbHide</strong>             | 0                      | Window is hidden and focus is passed to the hidden window. The  <strong>vbHide</strong> constant is not applicable on Macintosh platforms. |
| <strong>vbNormalFocus</strong>      | 1                      | Window has focus and is restored to its original size and position.                                                                        |
| <strong>vbMinimizedFocus</strong>   | 2                      | Window is displayed as an icon with focus.                                                                                                 |
| <strong>vbMaximizedFocus</strong>   | 3                      | Window is maximized with focus.                                                                                                            |
| <strong>vbNormalNoFocus</strong>    | 4                      | Window is restored to its most recent size and position. The currently active window remains active.                                       |
| <strong>vbMinimizedNoFocus</strong> | 6                      | Window is displayed as an icon. The currently active window remains active.                                                                |

 **Remarks**
If the  **Shell** function successfully executes the named file, it returns the task ID of the started program. The task ID is a unique number that identifies the running program. If the **Shell** function can't start the named program, an error occurs.
On the Macintosh,  **vbNormalFocus**, **vbMinimizedFocus**, and **vbMaximizedFocus** all place the application in the foreground; **vbHide**, **vbNoFocus**, **vbMinimizeFocus** all place the application in the background.

 **Note**  By default, the  **Shell** function runs other programs asynchronously. This means that a program started with **Shell** might not finish executing before the statements following the **Shell** function are executed.


## Example

This example uses the  **Shell** function to run an application specified by the user. On the MacIntosh, the default drive name is "HD" and portions of the pathname are separated by colons instead of backslashes. Similarly, you would specify Macintosh folders instead of \Windows.


```vb
' Specifying 1 as the second argument opens the application in 
' normal size and gives it the focus.
Dim RetVal
RetVal = Shell("C:\WINDOWS\CALC.EXE", 1)    ' Run Calculator.
```


