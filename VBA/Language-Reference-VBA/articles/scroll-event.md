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
 **Syntax**
For ScrollBar **Private Sub**_object_ _**Scroll( )**
For MultiPage **Private Sub**_object_ _**Scroll(**_index_**As Long**, _ActionX_**As fmScrollAction**, _ActionY_**As fmScrollAction**, **ByVal**_RequestDx_**As Single**, **ByVal**_RequestDy_**As Single**, **ByVal**_ActualDx_**As MSForms.ReturnSingle**, **ByVal**_ActualDy_**As MSForms.ReturnSingle)**
For Frame **Private Sub**_object_ _**Scroll(**_ActionX_**As fmScrollAction**, _ActionY_**As fmScrollAction**, **ByVal**_RequestDx_**As Single**, **ByVal**_RequestDy_**As Single**, **ByVal**_ActualDx_**As MSForms.ReturnSingle**, **ByVal**_ActualDy_**As MSForms.ReturnSingle)**
The  **Scroll** event syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object name.|
| _index_|Required. The index of the page in a  **MultiPage** associated with this event.|
| _ActionX_|Required. The action that occurred in the horizontal direction.|
| _ActionY_|Required. The action that occurred in the vertical direction.|
| _RequestDx_|Required. The distance, in points, you want the scroll bar to move in the horizontal direction.|
| _RequestDy_|Required. The distance, in points, you want the scroll bar to move in the vertical direction.|
| _ActualDx_|Required. The distance, in points, the scroll bar travelled in the horizontal direction.|
| _ActualDy_|Required. The distance, in points, the scroll bar travelled in the vertical direction.|
 **Settings**
The settings for  _ActionX_ and _ActionY_ are:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| _fmScrollActionNoChange_|0|No change occurred.|
| _fmScrollActionLineUp_|1|A small distance up on a vertical scroll bar; a small distance to the left on a horizontal scroll bar. Movement is equivalent to pressing the up or left arrow keys on the keyboard to move the scroll bar.|
| _fmScrollActionLineDown_|2|A small distance down on a vertical scroll bar; a small distance to the right on a horizontal scroll bar. Movement is equivalent to pressing the down or right arrow keys on the keyboard to move the scroll bar.|
| _fmScrollActionPageUp_|3|One page up on a vertical scroll bar; one page to the left on a horizontal scroll bar. Movement is equivalent to pressing PAGE UP on the keyboard to move the scroll bar.|
| _fmScrollActionPageDown_|4|One page down on a vertical scroll bar; one page to the right on a horizontal scroll bar. Movement is equivalent to pressing PAGE DOWN on the keyboard to move the scroll bar.|
| _fmScrollActionBegin_|5|The top of a vertical scroll bar; the left end of a horizontal scroll bar.|
| _fmScrollActionEnd_|6|The bottom of a vertical scroll bar; the right end of a horizontal scroll bar.|
| _fmScrollActionPropertyChange_|8|The value of either the  **ScrollTop** or the **ScrollLeft** property changed. The direction and amount of movement depend on which property was changed and on the new property value.|
| _fmScrollActionControlRequest_|9|A control asked its container to scroll. The amount of movement depends on the specific control and container involved.|
| _fmScrollActionFocusRequest_|10|The user moved to a different control. The amount of movement depends on the placement of the selected control, and generally has the effect of moving the selected control so it is completely visible to the user.|
 **Remarks**
The Scroll events associated with a form,  **Frame**, or **Page** return the following arguments: _ActionX_, _ActionY_, _ActualX_, and _ActualY_. _ActionX_ and _ActionY_ identify the action that occurred. _ActualX_ and _ActualY_ identify the distance that the scroll box traveled.
The default action is to calculate the new position of the scroll box and then scroll to that position.
You can initiate a Scroll event by issuing a  **Scroll** method for a form, **Frame**, or **Page**. Users can generate Scroll events by moving the scroll box.
The Scroll event associated with the stand-alone  **ScrollBar** indicates that the user moved the scroll box in either direction. This event is not initiated when the value of the **ScrollBar** changes by code or by the user clicking on parts of the **ScrollBar** other than the scroll box.

