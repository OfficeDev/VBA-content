---
title: DoEvents Function
keywords: vblr6.chm1014016
f1_keywords:
- vblr6.chm1014016
ms.prod: office
ms.assetid: b38afdfe-9f8a-ac15-3e02-47184dae69c5
ms.date: 06/08/2017
---


# DoEvents Function



Yields execution so that the operating system can process other events.
 **Syntax**
 **DoEvents(** **)**
 **Remarks**
The  **DoEvents** function returns an[Integer](vbe-glossary.md) representing the number of open forms in stand-alone versions of Visual Basic, such as Visual Basic, Professional Edition. **DoEvents** returns zero in all other applications.
 **DoEvents** passes control to the operating system. Control is returned after the operating system has finished processing the events in its queue and all keys in the **SendKeys** queue have been sent.
 **DoEvents** is most useful for simple things like allowing a user to cancel a process after it has started, for example a search for a file. For long-running processes, yielding the processor is better accomplished by using a Timer or delegating the task to an ActiveX EXE component. In the latter case, the task can continue completely independent of your application, and the operating system takes care of multitasking and time slicing.
Any time you temporarily yield the processor within an event procedure, make sure the [procedure](vbe-glossary.md) is not executed again from a different part of your code before the first call returns; this could cause unpredictable results. In addition, do not use **DoEvents** if other applications could possibly interact with your procedure in unforeseen ways during the time you have yielded control.

## Example

This example uses the  **DoEvents** function to cause execution to yield to the operating system once every 1000 iterations of the loop. **DoEvents** returns the number of open Visual Basic forms, but only when the host application is Visual Basic.


```vb
' Create a variable to hold number of Visual Basic forms loaded 
' and visible.
Dim I, OpenForms
For I = 1 To 150000    ' Start loop.
    If I Mod 1000 = 0 Then     ' If loop has repeated 1000 times.
        OpenForms = DoEvents    ' Yield to operating system.
    End If
Next I    ' Increment loop counter.


```


