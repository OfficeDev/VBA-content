---
title: Label.MouseDown Event (Access)
keywords: vbaac10.chm14286
f1_keywords:
- vbaac10.chm14286
ms.prod: access
api_name:
- Access.Label.MouseDown
ms.assetid: 1895790f-e5d6-493e-6005-a2c6a83cbd87
ms.date: 06/08/2017
---


# Label.MouseDown Event (Access)

The  **MouseDown** event occurs when the user presses a mouse button.


## Syntax

 _expression_. **MouseDown**( ** _Button_**, ** _Shift_**, ** _X_**, ** _Y_** )

 _expression_ A variable that represents a **Label** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Button_|Required|**Integer**|The button that was pressed to trigger the event. If you need to test for the Button argument, you can use one of the following intrinsic constants as bit masks:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p><b>acLeftButton</b>  The bit mask for the left mouse button.  
  </p></li><li><p><b>acRightButton</b>  The bit mask for the right mouse button.</p></li><li><p><b>acMiddleButton</b>  The bit mask for the middle mouse button.  
</p></li></ul>|
| _Shift_|Required|**Integer**|The state of the SHIFT, CTRL, and ALT keys when the button specified by the Button argument was pressed or released. If you need to test for the Shift argument, you can use one of the following intrinsic constants as bit masks:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p><b>acShiftMask</b>  The bit mask for the SHIFT key.  
  </p></li><li><p><b>acCtrlMask</b>  The bit mask for the CTRL key.</p></li><li><p><b>acAltMask</b>  The bit mask for the ALT key.  
  
 </p></li></ul>|
| _X_|Required|**Single**|The x coordinate for the current location of the mouse pointer, in twips. |
| _Y_|Required|**Single**|The y coordinate for the current location of the mouse pointer, in twips. |

## Remarks




- The  **MouseDown** event applies only to forms, form sections , and controls on a form, not controls on a report.
    
- This event does not apply to a label attached to another control, such as the label for a text box. It applies only to "freestanding" labels. Pressing and releasing a mouse button in an attached label has the same effect as pressing and releasing the button in the associated control. The normal events for the control occur; no separate events occur for the attached label.
    


To run a macro or event procedure when these events occur, set the  **OnMouseDown** property to the name of the macro or to [Event Procedure].

You can use a MouseDown event to specify what happens when a particular mouse button is pressed or released. Unlike the  **Click** and **DblClick** events, the MouseDown event enables you to distinguish between the left, right, and middle mouse buttons. You can also write code for mouse-keyboard combinations that use the SHIFT, CTRL, and ALT keys.

To cause a  **MouseDown** event for a form to occur, press the mouse button in a blank area or record selector on the form. To cause a MouseDown event for a form section to occur, press the mouse button in a blank area of the form section.

The following apply to  **MouseDown** events:




- If a mouse button is pressed while the pointer is over a form or control, that object receives all mouse events up to and including the last  **MouseUp** event.
    
- If mouse buttons are pressed in succession, the object that receives the mouse event after the first press receives all mouse events until all buttons are released.
    


To respond to an event caused by moving the mouse, you use a  **MouseMove** event.


## Example

The following example shows how you can find out which mouse button caused a MouseDown event.

To try the example, add the following event procedure to a form:




```vb
Private Sub Form_MouseDown(Button As Integer, _ 
         Shift As Integer, X As Single, _ 
         Y As Single) 
    If Button = acLeftButton Then 
        MsgBox "You pressed the left button." 
    End If 
    If Button = acRightButton Then 
        MsgBox "You pressed the right button." 
    End If 
    If Button = acMiddleButton Then 
        MsgBox "You pressed the middle button." 
    End If 
End Sub
```


## See also


#### Concepts


[Label Object](label-object-access.md)

