---
title: Application-defined or object-defined error
keywords: vblr6.chm1011299
f1_keywords:
- vblr6.chm1011299
ms.prod: office
ms.assetid: 4c1ea0e8-e6f6-a960-eb13-b4dfc2bf96fe
ms.date: 06/08/2017
---


# Application-defined or object-defined error

This message is displayed when an error generated with the  **Raise** method or **Error** statement doesn't correspond to an error defined by Visual Basic for Applications. It is also returned by the **Error** function for [arguments](vbe-glossary.md) that don't correspond to errors defined by Visual Basic for Applications. Thus it may be an error you defined, or one that is defined by an object, including [host applications](vbe-glossary.md) like Microsoft Excel, Visual Basic, and so on. For example, Visual Basic forms generate form-related errors that can't be generated from code simply by specifying a number as an argument to the **Raise** method or **Error** statement. This message has the following causes and solutions:



- Your application executed an  **Err.Raise**_n_ or **Error**_n_ statement, but the number _n_ isn't defined by Visual Basic for Applications. If this was what was intended, you must use **Err.Raise** and specify additional arguments so that an end user can understand the nature of the error. For example, you can include a description string, source, and help information. To regenerate an error that you trapped, this approach will work if you don't execute **Err.Clear** before regenerating the error. If you execute **Err.Clear** first, you must fill in the additional arguments to the **Raise** method. Look at the context in which the error occurred, and make sure you are regenerating the same error.
    
- It may be that in accessing objects from other applications, an error was propagated back to your program that can't be mapped to a Visual Basic error.

Check the documentation for any objects you have accessed. The  **Err** object's **Source** property should contain the programmatic ID of the application or object that generated the error. To understand the context of an error returned by an object, you may want to use the **On Error Resume Next** construct in code that accesses objects, rather than the **On Error GoTo**_line_ syntax.
    
## List trappable errors for the host application 
In the past, programmers often used a loop to print out a list of all trappable error message strings. Typically this was done with code such as the following:


```vb
For index = 1 to 500
    Debug.Print Error$(index)
Next index
```


Such code still lists all the Visual Basic for Applications error messages, but displays "Application-defined or object-defined error" for host-defined errors, for example those in Visual Basic that relate to forms, controls, and so on. Many of these are trappable [run-time errors](vbe-glossary.md). You can use the Help  **Search** dialog box to find the list of trappable errors specific to your host application. Click **Search**, type **Trappable** in the first text box, and then click **Show Topics**. Select **Trappable Errors** in the lower list box and click **Go To**.

For additional information, select the item in question and press F1 (in Windows) or HELP (on the Macintosh).

