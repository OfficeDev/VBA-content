---
title: NavigationControl.MouseMove Event (Access)
keywords: vbaac10.chm14204
f1_keywords:
- vbaac10.chm14204
ms.prod: access
api_name:
- Access.NavigationControl.MouseMove
ms.assetid: a5676866-db8b-078d-70dc-ee159c66671c
ms.date: 06/08/2017
---


# NavigationControl.MouseMove Event (Access)

The  **MouseMove** event occurs when the user moves the mouse.


## Syntax

 _expression_. **MouseMove**( ** _Button_**, ** _Shift_**, ** _X_**, ** _Y_** )

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




- The  **MouseMove** event applies only to forms, form sections , and controls on a form, not controls on a report.
    
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


[NavigationControl Object](navigationcontrol-object-access.md)

