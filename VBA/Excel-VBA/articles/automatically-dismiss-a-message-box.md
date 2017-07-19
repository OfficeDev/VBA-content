---
title: Automatically Dismiss a Message Box
ms.prod: excel
ms.assetid: e4a38fbe-6bed-45dd-98cd-d10376f84322
ms.date: 06/08/2017
---


# Automatically Dismiss a Message Box

This example shows how to automatically dismiss a message box after a specified period of time. This example displays a message box and then automatically dismisses it after 10 seconds.

 **Sample code provided by:** Tom Urtis, [Atlas Programming Management](http://www.atlaspm.com/)



```vb
Sub MessageBoxTimer()
    Dim AckTime As Integer, InfoBox As Object
    Set InfoBox = CreateObject("WScript.Shell")
    'Set the message box to close after 10 seconds
    AckTime = 10
    Select Case InfoBox.Popup("Click OK (this window closes automatically after 10 seconds).", _
    AckTime, "This is your Message Box", 0)
        Case 1, -1
            Exit Sub
    End Select
End Sub
```


## About the Contributor
<a name="AboutContributor"> </a>

MVP Tom Urtis is the founder of Atlas Programming Management, a full-service Microsoft Office and Excel business solutions company in Silicon Valley. Tom has over 25 years of experience in business management and developing Microsoft Office applications, and is the coauthor of "Holy Macro! It's 2,500 Excel VBA Examples." 


