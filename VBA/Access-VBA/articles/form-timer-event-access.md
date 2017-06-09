---
title: Form.Timer Event (Access)
keywords: vbaac10.chm13659
f1_keywords:
- vbaac10.chm13659
ms.prod: access
api_name:
- Access.Form.Timer
ms.assetid: 395c62a1-5731-01b8-a4ea-852bfb30572f
ms.date: 06/08/2017
---


# Form.Timer Event (Access)

The  **Timer** event occurs for a form at regular intervals as specified by the form's **[TimerInterval](form-timerinterval-property-access.md)** property.


## Syntax

 _expression_. **Timer**

 _expression_ A variable that represents a **Form** object.


## Remarks

To run a macro or event procedure when this event occurs, set the  **OnTimer** property to the name of the macro or to [Event Procedure].

By running a macro or event procedure when a  **Timer** event occurs, you can control what Microsoft Access does at every timer interval. For example, you might want to requery underlying records or repaint the screen at specified intervals.

The  **TimerInterval** property setting of the form specifies the interval, in milliseconds, between **Timer** events. The interval can be between 0 and 2,147,483,647 milliseconds. Setting the **TimerInterval** property to 0 prevents the **Timer** event from occurring.

 **Link provided by:**
![Community Member Icon](images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) The[UtterAccess](http://www.utteraccess.com) community


- [Delay Event/Actions for Set Time Interval](http://www.utteraccess.com/wiki/index.php/Delay_Event/Actions_for_Set_Time_Interval)
    

## Example

The following example demonstrates a digital clock you can display on a form. A label control displays the current time according to your computer's system clock. 

To try the example, add the following event procedure to a form that contains a label named Clock. Set the form's  **TimerInterval** property to 1000 milliseconds to update the clock every second.




```vb
Private Sub Form_Timer() 
    Clock.Caption = Time        ' Update time display. 
End Sub
```


## About the Contributors
<a name="AboutContributors"> </a>

UtterAccess is the premier Microsoft Access wiki and help forum. Click here to join. 


## See also
<a name="AboutContributors"> </a>


#### Concepts


[Form Object](form-object-access.md)

