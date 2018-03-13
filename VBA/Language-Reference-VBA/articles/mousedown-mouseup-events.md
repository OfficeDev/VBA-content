---
title: MouseDown, MouseUp Events
keywords: fm20.chm5224947
f1_keywords:
- fm20.chm5224947
ms.prod: office
ms.assetid: 760c2492-4a33-8d17-eeef-e52da662d4c4
ms.date: 06/08/2017
---


# MouseDown, MouseUp Events



Occur when the user clicks a mouse button. MouseDown occurs when the user presses the mouse button; MouseUp occurs when the user releases the mouse button.
 <strong>Syntax</strong>
For MultiPage, TabStrip 
<strong>Private Sub</strong><em>object</em> <em><strong>MouseDown(</strong>_index</em><strong>As Long</strong>, <strong>ByVal</strong><em>Button</em><strong>As fmButton</strong>, <strong>ByVal</strong><em>Shift</em><strong>As fmShiftState</strong>, <strong>ByVal</strong><em>X</em><strong>As Single</strong>, <strong>ByVal</strong><em>Y</em><strong>As Single)</strong> <strong>Private Sub</strong><em>object</em> <em><strong>MouseUp(</strong>_index</em><strong>As Long</strong>, <strong>ByVal</strong><em>Button</em><strong>As fmButton</strong>, <strong>ByVal</strong><em>Shift</em><strong>As fmShiftState</strong>, <strong>ByVal</strong><em>X</em><strong>As Single</strong>, <strong>ByVal</strong><em>Y</em><strong>As Single)</strong>
For other controls 
<strong>Private Sub</strong><em>object</em> <em><strong>MouseDown( ByVal</strong>_Button</em><strong>As fmButton</strong>, <strong>ByVal</strong><em>Shift</em><strong>As fmShiftState</strong>, <strong>ByVal</strong><em>X</em><strong>As Single</strong>, <strong>ByVal</strong><em>Y</em><strong>As Single)</strong> <strong>Private Sub</strong><em>object</em> <em><strong>MouseUp( ByVal</strong>_Button</em><strong>As fmButton</strong>, <strong>ByVal</strong><em>Shift</em><strong>As fmShiftState</strong>, <strong>ByVal</strong><em>X</em><strong>As Single</strong>, <strong>ByVal</strong><em>Y</em><strong>As Single)</strong>
The  
<strong>MouseDown</strong> and <strong>MouseUp</strong> event syntaxes have these parts:


| <strong>Part</strong> | <strong>Description</strong>                                                                                                                        |
|:----------------------|:----------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>object</em>       | Required. A valid object.                                                                                                                           |
| <em>index</em>        | Required. The index of the page or tab in a  <strong>MultiPage</strong> or <strong>TabStrip</strong> with the specified event.                      |
| <em>Button</em>       | Required. An integer value that identifies which mouse button caused the event.                                                                     |
| <em>Shift</em>        | Required. The state of SHIFT, CTRL, and ALT.                                                                                                        |
| <em>X, Y</em>         | Required. The horizontal or vertical position, in points, from the left or top edge of the form,  <strong>Frame</strong>, or <strong>Page</strong>. |

 **Settings**
The settings for  _Button_ are:


| <strong>Constant</strong> | <strong>Value</strong> | <strong>Description</strong>   |
|:--------------------------|:-----------------------|:-------------------------------|
| <em>fmButtonLeft</em>     | 1                      | The left button was pressed.   |
| <em>fmButtonRight</em>    | 2                      | The right button was pressed.  |
| <em>fmButtonMiddle</em>   | 4                      | The middle button was pressed. |

The settings for  _Shift_ are:


| <strong>Value</strong> | <strong>Description</strong>       |
|:-----------------------|:-----------------------------------|
| 1                      | SHIFT was pressed.                 |
| 2                      | CTRL was pressed.                  |
| 3                      | SHIFT and CTRL were pressed.       |
| 4                      | ALT was pressed.                   |
| 5                      | ALT and SHIFT were pressed.        |
| 6                      | ALT and CTRL were pressed.         |
| 7                      | ALT, SHIFT, and CTRL were pressed. |

You can identify individual keyboard modifiers by using the following constants:


| <strong>Constant</strong> | <strong>Value</strong> | <strong>Description</strong> |
|:--------------------------|:-----------------------|:-----------------------------|
| <em>fmShiftMask</em>      | 1                      | Mask to detect SHIFT.        |
| <em>fmCtrlMask</em>       | 2                      | Mask to detect CTRL.         |
| <em>fmAltMask</em>        | 4                      | Mask to detect ALT.          |

 **Remarks**
For a  **MultiPage**, the MouseDown event occurs when the user presses a mouse button over the control.
For a  **TabStrip**, the index argument identifies the tab where the user clicked. An index of -1 indicates the user did not click a tab. For example, if there are no tabs in the upper right corner of the control, clicking in the upper right corner sets the index to -1.
For a form, the user can generate MouseDown and MouseUp events by pressing and releasing a mouse button in a blank area, record selector, or scroll bar on the form.
The sequence of mouse-related events is:


1. MouseDown

2. MouseUp

3. Click

4. DblClick

5. MouseUp


MouseDown or MouseUp event procedures specify actions that occur when a mouse button is pressed or released. MouseDown and MouseUp events enable you to distinguish between the left, right, and middle mouse buttons. You can also write code for mouse-keyboard combinations that use the SHIFT, CTRL, and ALT keyboard modifiers.
If a mouse button is pressed while the pointer is over a form or control, that object "captures" the mouse and receives all mouse events up to and including the last MouseUp event. This implies that the  _X_, _Y_ mouse-pointer coordinates returned by a mouse event may not always be within the boundaries of the object that receives them.
If mouse buttons are pressed in succession, the object that captures the mouse receives all successive mouse events until all buttons are released.
Use the  _Shift_ argument to identify the state of SHIFT, CTRL, and ALT when the MouseDown or MouseUp event occurred. For example, if both CTRL and ALT are pressed, the value of _Shift_ is 6.

