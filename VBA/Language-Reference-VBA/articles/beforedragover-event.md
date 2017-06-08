---
title: BeforeDragOver Event
keywords: fm20.chm5224935
f1_keywords:
- fm20.chm5224935
ms.prod: office
api_name:
- Office.BeforeDragOver
ms.assetid: 0c2803fc-0f69-60d8-06fb-36870aad8a27
ms.date: 06/08/2017
---


# BeforeDragOver Event



Occurs when a drag-and-drop operation is in progress.
 **Syntax**
For Frame **Private Sub**_object_ _**BeforeDragOver( ByVal**_Cancel_**As MSForms.ReturnBoolean**, _ctrl_**As Control**, **ByVal**_Data_**As DataObject**, **ByVal**_X_**As Single**, **ByVal**_Y_**As Single**, **ByVal**_DragState_**As fmDragState**, **ByVal**_Effect_**As MSForms.ReturnEffect**, **ByVal**_Shift_**As fmShiftState)**
For MultiPage **Private Sub**_object_ _**BeforeDragOver(**_index_**As Long**, **ByVal**_Cancel_**As MSForms.ReturnBoolean**, _ctrl_**As Control**, **ByVal**_Data_**As DataObject**, **ByVal**_X_**As Single**, **ByVal**_Y_**As Single**, **ByVal**_DragState_**As fmDragState**, **ByVal**_Effect_**As MSForms.ReturnEffect**, **ByVal**_Shift_**As fmShiftState)**
For TabStrip **Private Sub**_object_ _**BeforeDragOver(**_index_**As Long**, **ByVal**_Cancel_**As MSForms.ReturnBoolean**, **ByVal**_Data_**As DataObject**, **ByVal**_X_**As Single**, **ByVal**_Y_**As Single**, **ByVal**_DragState_**As fmDragState**, **ByVal**_Effect_**As MSForms.ReturnEffect**, **ByVal**_Shift_**As fmShiftState)**
For other controls **Private Sub**_object_ _**BeforeDragOver( ByVal**_Cancel_**As MSForms.ReturnBoolean**, **ByVal**_Data_**As DataObject**, **ByVal**_X_**As Single**, **ByVal**_Y_**As Single**, **ByVal**_DragState_**As fmDragState**, **ByVal**_Effect_**As MSForms.ReturnEffect**, **ByVal**_Shift_**As fmShiftState)**
The  **BeforeDragOver** event syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object name.|
| _index_|Required. The index of the  **Page** in a **MultiPage** that the drag-and-drop operation will affect.|
| _Cancel_|Required. Event status.  **False** indicates that the control should handle the event (default). **True** indicates the application handles the event.|
| _ctrl_|Required. The control being dragged over.|
| _Data_|Required. Data that is dragged in a drag-and-drop operation. The data is packaged in a  **DataObject**.|
| _X, Y_|Required. The horizontal and vertical coordinates of the control's position. Both coordinates are measured in points.  _X_ is measured from the left edge of the control; _Y_ is measured from the top of the control..|
| _DragState_|Required. Transition state of the data being dragged.|
| _Effect_|Required. Operations supported by the [drop source](glossary-vba.md).|
| _Shift_|Required. Specifies the state of SHIFT, CTRL, and ALT.|
 **Settings**
The settings for  _DragState_ are:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| _fmDragStateEnter_|0|Mouse pointer is within range of a target.|
| _fmDragStateLeave_|1|Mouse pointer is outside the range of a target.|
| _fmDragStateOver_|2|Mouse pointer is at a new position, but remains within range of the same target.|
The settings for  _Effect_ are:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| _fmDropEffectNone_|0|Does not copy or move the drop source to the drop target.|
| _fmDropEffectCopy_|1|Copies the drop source to the drop target.|
| _fmDropEffectMove_|2|Moves the drop source to the drop target.|
| _fmDropEffectCopyOrMove_|3|Copies or moves the drop source to the drop target.|
The settings for  _Shift_ are:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| _fmShiftMask_|1|SHIFT was pressed.|
| _fmCtrlMask_|2|CTRL was pressed.|
| _fmAltMask_|4|ALT was pressed.|
 **Remarks**
Use this event to monitor the mouse pointer as it enters, leaves, or rests directly over a valid [target](glossary-vba.md). When a drag-and-drop operation is in progress, the system initiates this event when the user moves the mouse, or presses or releases the mouse button or buttons. The mouse pointer position determines the target object that receives this event. You can determine the state of the mouse pointer by examining the  _DragState_ argument.
When a control handles this event, you can use the  _Effect_ argument to identify the drag-and-drop action to perform. When _Effect_ is set to **fmDropEffectCopyOrMove**, the drop source supports a copy ( **fmDropEffectCopy** ), move ( **fmDropEffectMove** ), or a cancel ( **fmDropEffectNone** ) operation.
When  _Effect_ is set to **fmDropEffectCopy**, the drop source supports a copy or a cancel ( **fmDropEffectNone** ) operation.
When  _Effect_ is set to **fmDropEffectMove**, the drop source supports a move or a cancel ( **fmDropEffectNone** ) operation.
When  _Effect_ is set to **fmDropEffectNone**. the drop source supports a cancel operation.
Most controls do not support drag-and-drop while  _Cancel_ is **False**, which is the default setting. This means the control rejects attempts to drag or drop anything on the control, and the control does not initiate the BeforeDropOrPaste event. The **TextBox** and **ComboBox** controls are exceptions to this; these controls support drag-and-drop operations even when _Cancel_ is **False**.

