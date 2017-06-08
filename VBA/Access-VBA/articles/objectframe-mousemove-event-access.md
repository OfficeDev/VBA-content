---
title: ObjectFrame.MouseMove Event (Access)
keywords: vbaac10.chm14112
f1_keywords:
- vbaac10.chm14112
ms.prod: access
api_name:
- Access.ObjectFrame.MouseMove
ms.assetid: aad412fd-96ba-803d-848c-a8788fa1e0ae
ms.date: 06/08/2017
---


# ObjectFrame.MouseMove Event (Access)

The  **MouseMove** event occurs when the user moves the mouse.


## Syntax

 _expression_. **MouseMove**( ** _Button_**, ** _Shift_**, ** _X_**, ** _Y_** )

 _expression_ A variable that represents an **ObjectFrame** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Button_|Required|**Integer**|The button that was pressed or released when the event was triggered. If you need to test for the Button argument, you can use one of the following intrinsic constants as bit masks:
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




- The  **MouseMove** event applies only to forms, form sections, and controls on a form, not controls on a report.
    
- This event does not apply to a label attached to another control, such as the label for a text box. It applies only to "freestanding" labels. Pressing and releasing a mouse button in an attached label has the same effect as pressing and releasing the button in the associated control. The normal events for the control occur; no separate events occur for the attached label.
    


To run a macro or event procedure when these events occur, set the  **OnMouseMove** property to the name of the macro or to [Event Procedure].

The  **MouseMove** event is generated continually as the mouse pointer moves over objects. Unless another object generates a mouse event, an object recognizes a MouseMove event whenever the mouse pointer is positioned within its borders.

To cause a  **MouseMove** event for a form to occur, move the mouse pointer over a blank area, record selector, or scroll bar on the form. To cause a **MouseMove** event for a form section to occur, move the mouse pointer over a blank area of the form section.

To respond to an event caused by moving the mouse, you use a  **MouseMove** event.


 **Note**  

To run a macro or event procedure in response to pressing and releasing the mouse buttons, you use the  **MouseDown** and **MouseUp** events.


## Example

The following example determines where the mouse is and whether the left mouse button and/or the SHIFT key is pressed. The x and y coordinates of the mouse pointer position are displayed in a label control as you move the mouse.


```vb
Private Sub Detail_MouseMove(Button As Integer, _ 
     Shift As Integer, X As Single, Y As Single) 
    Dim intShiftDown As Integer, intLeftButton As Integer 
 
    Me!Coordinates.Caption = X &; ", " &; Y 
    ' Use bit masks to determine state of 
    ' SHIFT key and left button. 
    intShiftDown = Shift And acShiftMask 
    intLeftButton = Button And acLeftButton 
    ' Check that SHIFT key and left button  
    ' are both pressed. 
    If intShiftDown And intLeftButton > 0 Then 
        MsgBox "Shift key and left mouse button were pressed." 
    End If 
End Sub
```


## See also


#### Concepts


[ObjectFrame Object](objectframe-object-access.md)

