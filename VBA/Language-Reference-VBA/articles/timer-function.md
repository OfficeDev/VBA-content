---
title: Timer Function
keywords: vblr6.chm1009043
f1_keywords:
- vblr6.chm1009043
ms.prod: office
ms.assetid: a39cf81a-a90c-5833-75e8-9ac4605e3b02
ms.date: 06/08/2017
---


# Timer Function



Returns a  **Single** representing the number of seconds elapsed since midnight.
 **Syntax**
 **Timer**
 **Remarks**
In Microsoft Windows the  **Timer** function returns fractional portions of a second. On the Macintosh, timer resolution is one second.

## Example

This example uses the  **Timer** function to pause the application. The example also uses **DoEvents** to yield to other processes during the pause.


```vb
Dim PauseTime, Start, Finish, TotalTime
If (MsgBox("Press Yes to pause for 5 seconds", 4)) = vbYes Then
    PauseTime = 5    ' Set duration.
    Start = Timer    ' Set start time.
    Do While Timer < Start + PauseTime
        DoEvents    ' Yield to other processes.
    Loop
    Finish = Timer    ' Set end time.
    TotalTime = Finish - Start    ' Calculate total time.
    MsgBox "Paused for " &; TotalTime &; " seconds"
Else
    End
End If

```


