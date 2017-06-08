---
title: NavigationControl.MouseUp Event (Access)
keywords: vbaac10.chm14205
f1_keywords:
- vbaac10.chm14205
ms.prod: access
api_name:
- Access.NavigationControl.MouseUp
ms.assetid: 174c4b0d-9906-5f73-80a2-a59b3d66aae1
ms.date: 06/08/2017
---


# NavigationControl.MouseUp Event (Access)

The  **MouseUp** event occurs when the user releases a mouse button.


## Syntax

 _expression_. **MouseUp**( ** _Button_**, ** _Shift_**, ** _X_**, ** _Y_** )

 _expression_ A variable that represents a **NavigationControl** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Button_|Required|**Integer**||
| _Shift_|Required|**Integer**||
| _X_|Required|**Single**||
| _Y_|Required|**Single**||

### Return Value

nothing


## Remarks




- The  **MouseUp** event applies only to forms, form sections , and controls on a form, not controls on a report.
    
- This event does not apply to a label attached to another control, such as the label for a text box. It applies only to "freestanding" labels. Pressing and releasing a mouse button in an attached label has the same effect as pressing and releasing the button in the associated control. The normal events for the control occur; no separate events occur for the attached label.
    


To run a macro or event procedure when these events occur, set the  **OnMouseUp** property to the name of the macro or to [Event Procedure].

You can use a  **MouseUp** event to specify what happens when a particular mouse button is pressed or released. Unlike the **Click** and **DblClick** events, the **MouseUp** event enables you to distinguish between the left, right, and middle mouse buttons. You can also write code for mouse-keyboard combinations that use the SHIFT, CTRL, and ALT keys.

To cause a  **MouseUp** event for a form to occur, press the mouse button in a blank area or record selector on the form. To cause a **MouseUp** event for a form section to occur, press the mouse button in a blank area of the form section.

The following apply to  **MouseUp** events:




- If a mouse button is pressed while the pointer is over a form or control, that object receives all mouse events up to and including the last  **MouseUp** event.
    
- If mouse buttons are pressed in succession, the object that receives the mouse event after the first press receives all mouse events until all buttons are released.
    


To respond to an event caused by moving the mouse, you use a  **MouseMove** event.


## See also


#### Concepts


[NavigationControl Object](navigationcontrol-object-access.md)

