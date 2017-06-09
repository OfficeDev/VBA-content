---
title: Report.TimerInterval Property (Access)
keywords: vbaac10.chm13825
f1_keywords:
- vbaac10.chm13825
ms.prod: access
api_name:
- Access.Report.TimerInterval
ms.assetid: 272fb1f6-2aca-60c2-1f0f-d901e0da91ac
ms.date: 06/08/2017
---


# Report.TimerInterval Property (Access)

You can use the  **TimerInterval** property to specify the interval, in milliseconds, between **[Timer](report-timer-event-access.md)** events on a report. Read/write **Long**.


## Syntax

 _expression_. **TimerInterval**

 _expression_ A variable that represents a **Report** object.


## Remarks

The  **TimerInterval** property setting is a Long Integer value between 0 and 2,147,483,647.

You can set this property by using the report's property sheet, a macro, or Visual Basic.


 **Note**  When using Visual Basic, you set the  **TimerInterval** property in the report's **Load** event.

To run Visual Basic code at intervals specified by the  **TimerInterval** property, put the code in the report's **Timer** event procedure. For example, to requery records every 30 seconds, put the code to requery the records in the report's **Timer** event procedure, and then set the **TimerInterval** property to 30000.


## Example

The following example shows how to create a flashing button on a form by displaying and hiding an icon on the button. The report's  **Load** event procedure sets the report's **TimerInterval** property to 1000 so the icon display is toggled once every second.


```vb
Sub Report_Load() 
 Me.TimerInterval = 1000 
End Sub 
 
Sub Report_Timer() 
 Static intShowPicture As Integer 
 If intShowPicture Then 
 ' Show icon. 
 Me!btnPicture.Picture = "C:\Icons\Flash.ico" 
 Else 
 ' Don't show icon. 
 Me!btnPicture.Picture = "" 
 End If 
 intShowPicture = Not intShowPicture 
End Sub
```


## See also


#### Concepts


[Report Object](report-object-access.md)

