---
title: Clear Method (Visual Basic for Applications)
keywords: vblr6.chm1014184
f1_keywords:
- vblr6.chm1014184
ms.prod: office
ms.assetid: 90766255-52c5-a230-b8aa-c66302f452d2
ms.date: 06/08/2017
---


# Clear Method (Visual Basic for Applications)



Clears all [property](vbe-glossary.md) settings of the **Err** object.
 **Syntax**
 _object_**.Clear**
The  _object_ is always the **Err** object.
 **Remarks**
Use  **Clear** to explicitly clear the **Err** object after an error has been handled, for example, when you use deferred error handling with **On Error Resume Next**. The **Clear** method is called automatically whenever any of the following[statements](vbe-glossary.md) is executed:


- Any type of  **Resume** statement
    
-  **Exit Sub**, **Exit Function**, **Exit Property**
    
- Any  **On Error** statement
    
     **Note**  The  **On Error Resume Next** construct may be preferable to **On Error GoTo** when handling errors generated during access to other objects. Checking **Err** after each interaction with an object removes ambiguity about which object was accessed by the code. You can be sure which object placed the error code in **Err.Number**, as well as which object originally generated the error (the object specified in **Err.Source** ).


## Example

This example uses the  **Err** object's **Clear** method to reset the numeric properties of the **Err** object to zero and its string properties to zero-length strings. If **Clear** were omitted from the following code, the error message dialog box would be displayed on every iteration of the loop (after an error occurs) whether or not a successive calculation generated an error. You can single-step through the code to see the effect.


```vb
Dim Result(10) As Integer    ' Declare array whose elements 
            ' will overflow easily.
Dim indx
On Error Resume Next    ' Defer error trapping.
Do Until indx = 10
    ' Generate an occasional error or store result if no error.
    Result(indx) = Rnd * indx * 20000
    If Err.Number <> 0 Then
        MsgBox Err, , "Error Generated: ", Err.HelpFile, Err.HelpContext
        Err.Clear    ' Clear Err object properties.
    End If
    indx = indx + 1
Loop

```


