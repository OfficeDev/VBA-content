---
title: Scroll Event
keywords: fm20.chm5224950
f1_keywords:
- fm20.chm5224950
ms.prod: office
api_name:
- Office.Scroll
ms.assetid: 1b4f6243-ea9b-320c-1afd-9bb230823ffb
ms.date: 06/08/2017
---


# Scroll Event



Occurs when the scroll box is repositioned.
 <strong>Syntax</strong>
For ScrollBar 
<strong>Private Sub</strong><em>object</em> <em><strong>Scroll( )</strong>
For MultiPage <strong>Private Sub</strong>_object</em> <em><strong>Scroll(</strong>_index</em><strong>As Long</strong>, <em>ActionX</em><strong>As fmScrollAction</strong>, <em>ActionY</em><strong>As fmScrollAction</strong>, <strong>ByVal</strong><em>RequestDx</em><strong>As Single</strong>, <strong>ByVal</strong><em>RequestDy</em><strong>As Single</strong>, <strong>ByVal</strong><em>ActualDx</em><strong>As MSForms.ReturnSingle</strong>, <strong>ByVal</strong><em>ActualDy</em><strong>As MSForms.ReturnSingle)</strong>
For Frame 
<strong>Private Sub</strong><em>object</em> <em><strong>Scroll(</strong>_ActionX</em><strong>As fmScrollAction</strong>, <em>ActionY</em><strong>As fmScrollAction</strong>, <strong>ByVal</strong><em>RequestDx</em><strong>As Single</strong>, <strong>ByVal</strong><em>RequestDy</em><strong>As Single</strong>, <strong>ByVal</strong><em>ActualDx</em><strong>As MSForms.ReturnSingle</strong>, <strong>ByVal</strong><em>ActualDy</em><strong>As MSForms.ReturnSingle)</strong>
The  
<strong>Scroll</strong> event syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong>                                                                    |
|:----------------------|:------------------------------------------------------------------------------------------------|
| <em>object</em>       | Required. A valid object name.                                                                  |
| <em>index</em>        | Required. The index of the page in a  <strong>MultiPage</strong> associated with this event.    |
| <em>ActionX</em>      | Required. The action that occurred in the horizontal direction.                                 |
| <em>ActionY</em>      | Required. The action that occurred in the vertical direction.                                   |
| <em>RequestDx</em>    | Required. The distance, in points, you want the scroll bar to move in the horizontal direction. |
| <em>RequestDy</em>    | Required. The distance, in points, you want the scroll bar to move in the vertical direction.   |
| <em>ActualDx</em>     | Required. The distance, in points, the scroll bar travelled in the horizontal direction.        |
| <em>ActualDy</em>     | Required. The distance, in points, the scroll bar travelled in the vertical direction.          |

 **Settings**
The settings for  _ActionX_ and _ActionY_ are:


| <strong>Constant</strong>             | <strong>Value</strong> | <strong>Description</strong>                                                                                                                                                                                         |
|:--------------------------------------|:-----------------------|:---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>fmScrollActionNoChange</em>       | 0                      | No change occurred.                                                                                                                                                                                                  |
| <em>fmScrollActionLineUp</em>         | 1                      | A small distance up on a vertical scroll bar; a small distance to the left on a horizontal scroll bar. Movement is equivalent to pressing the up or left arrow keys on the keyboard to move the scroll bar.          |
| <em>fmScrollActionLineDown</em>       | 2                      | A small distance down on a vertical scroll bar; a small distance to the right on a horizontal scroll bar. Movement is equivalent to pressing the down or right arrow keys on the keyboard to move the scroll bar.    |
| <em>fmScrollActionPageUp</em>         | 3                      | One page up on a vertical scroll bar; one page to the left on a horizontal scroll bar. Movement is equivalent to pressing PAGE UP on the keyboard to move the scroll bar.                                            |
| <em>fmScrollActionPageDown</em>       | 4                      | One page down on a vertical scroll bar; one page to the right on a horizontal scroll bar. Movement is equivalent to pressing PAGE DOWN on the keyboard to move the scroll bar.                                       |
| <em>fmScrollActionBegin</em>          | 5                      | The top of a vertical scroll bar; the left end of a horizontal scroll bar.                                                                                                                                           |
| <em>fmScrollActionEnd</em>            | 6                      | The bottom of a vertical scroll bar; the right end of a horizontal scroll bar.                                                                                                                                       |
| <em>fmScrollActionPropertyChange</em> | 8                      | The value of either the  <strong>ScrollTop</strong> or the <strong>ScrollLeft</strong> property changed. The direction and amount of movement depend on which property was changed and on the new property value.    |
| <em>fmScrollActionControlRequest</em> | 9                      | A control asked its container to scroll. The amount of movement depends on the specific control and container involved.                                                                                              |
| <em>fmScrollActionFocusRequest</em>   | 10                     | The user moved to a different control. The amount of movement depends on the placement of the selected control, and generally has the effect of moving the selected control so it is completely visible to the user. |

 **Remarks**
The Scroll events associated with a form,  **Frame**, or **Page** return the following arguments: _ActionX_, _ActionY_, _ActualX_, and _ActualY_. _ActionX_ and _ActionY_ identify the action that occurred. _ActualX_ and _ActualY_ identify the distance that the scroll box traveled.
The default action is to calculate the new position of the scroll box and then scroll to that position.
You can initiate a Scroll event by issuing a  **Scroll** method for a form, **Frame**, or **Page**. Users can generate Scroll events by moving the scroll box.
The Scroll event associated with the stand-alone  **ScrollBar** indicates that the user moved the scroll box in either direction. This event is not initiated when the value of the **ScrollBar** changes by code or by the user clicking on parts of the **ScrollBar** other than the scroll box.

