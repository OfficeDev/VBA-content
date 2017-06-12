---
title: Form.TimerInterval Property (Access)
keywords: vbaac10.chm13462
f1_keywords:
- vbaac10.chm13462
ms.prod: access
api_name:
- Access.Form.TimerInterval
ms.assetid: ee56bcf8-20cb-9d86-ed17-3b85ac88f6f1
ms.date: 06/08/2017
---


# Form.TimerInterval Property (Access)

You can use the  **TimerInterval** property to specify the interval, in milliseconds, between **[Timer](form-timer-event-access.md)** events on a form. Read/write Long.


## Syntax

 _expression_. **TimerInterval**

 _expression_ A variable that represents a **Form** object.


## Remarks

The  **TimerInterval** property setting is a Long Integer value between 0 and 2,147,483,647.

You can set this property by using the form's property sheet, a macro, or Visual Basic.


 **Note**  When using Visual Basic, you set the  **TimerInterval** property in the form's **Load** event.

To run Visual Basic code at intervals specified by the  **TimerInterval** property, put the code in the form's **Timer** event procedure. For example, to requery records every 30 seconds, put the code to requery the records in the form's **Timer** event procedure, and then set the **TimerInterval** property to 30000.

 **Link provided by:**
![Community Member Icon](images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) The[UtterAccess](http://www.utteraccess.com) community


- [Delay Event/Actions for Set Time Interval](http://www.utteraccess.com/wiki/index.php/Delay_Event/Actions_for_Set_Time_Interval)
    

## Example

The following example shows how to create a flashing button on a form by displaying and hiding an icon on the button. The form's  **Load** event procedure sets the form's **TimerInterval** property to 1000 so the icon display is toggled once every second.


```vb
Sub Form_Load() 
    Me.TimerInterval = 1000 
End Sub 
 
Sub Form_Timer() 
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


## About the Contributors
<a name="AboutContributors"> </a>

UtterAccess is the premier Microsoft Access wiki and help forum. Click here to join. 


## See also
<a name="AboutContributors"> </a>


#### Concepts


[Form Object](form-object-access.md)

