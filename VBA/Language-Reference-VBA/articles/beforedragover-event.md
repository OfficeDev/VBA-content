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
 <strong>Syntax</strong>
For Frame 
<strong>Private Sub</strong><em>object</em> <em><strong>BeforeDragOver( ByVal</strong>_Cancel</em><strong>As MSForms.ReturnBoolean</strong>, <em>ctrl</em><strong>As Control</strong>, <strong>ByVal</strong><em>Data</em><strong>As DataObject</strong>, <strong>ByVal</strong><em>X</em><strong>As Single</strong>, <strong>ByVal</strong><em>Y</em><strong>As Single</strong>, <strong>ByVal</strong><em>DragState</em><strong>As fmDragState</strong>, <strong>ByVal</strong><em>Effect</em><strong>As MSForms.ReturnEffect</strong>, <strong>ByVal</strong><em>Shift</em><strong>As fmShiftState)</strong>
For MultiPage 
<strong>Private Sub</strong><em>object</em> <em><strong>BeforeDragOver(</strong>_index</em><strong>As Long</strong>, <strong>ByVal</strong><em>Cancel</em><strong>As MSForms.ReturnBoolean</strong>, <em>ctrl</em><strong>As Control</strong>, <strong>ByVal</strong><em>Data</em><strong>As DataObject</strong>, <strong>ByVal</strong><em>X</em><strong>As Single</strong>, <strong>ByVal</strong><em>Y</em><strong>As Single</strong>, <strong>ByVal</strong><em>DragState</em><strong>As fmDragState</strong>, <strong>ByVal</strong><em>Effect</em><strong>As MSForms.ReturnEffect</strong>, <strong>ByVal</strong><em>Shift</em><strong>As fmShiftState)</strong>
For TabStrip 
<strong>Private Sub</strong><em>object</em> <em><strong>BeforeDragOver(</strong>_index</em><strong>As Long</strong>, <strong>ByVal</strong><em>Cancel</em><strong>As MSForms.ReturnBoolean</strong>, <strong>ByVal</strong><em>Data</em><strong>As DataObject</strong>, <strong>ByVal</strong><em>X</em><strong>As Single</strong>, <strong>ByVal</strong><em>Y</em><strong>As Single</strong>, <strong>ByVal</strong><em>DragState</em><strong>As fmDragState</strong>, <strong>ByVal</strong><em>Effect</em><strong>As MSForms.ReturnEffect</strong>, <strong>ByVal</strong><em>Shift</em><strong>As fmShiftState)</strong>
For other controls 
<strong>Private Sub</strong><em>object</em> <em><strong>BeforeDragOver( ByVal</strong>_Cancel</em><strong>As MSForms.ReturnBoolean</strong>, <strong>ByVal</strong><em>Data</em><strong>As DataObject</strong>, <strong>ByVal</strong><em>X</em><strong>As Single</strong>, <strong>ByVal</strong><em>Y</em><strong>As Single</strong>, <strong>ByVal</strong><em>DragState</em><strong>As fmDragState</strong>, <strong>ByVal</strong><em>Effect</em><strong>As MSForms.ReturnEffect</strong>, <strong>ByVal</strong><em>Shift</em><strong>As fmShiftState)</strong>
The  
<strong>BeforeDragOver</strong> event syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong>                                                                                                                                                                                                          |
|:----------------------|:--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>object</em>       | Required. A valid object name.                                                                                                                                                                                                        |
| <em>index</em>        | Required. The index of the  <strong>Page</strong> in a <strong>MultiPage</strong> that the drag-and-drop operation will affect.                                                                                                       |
| <em>Cancel</em>       | Required. Event status.  <strong>False</strong> indicates that the control should handle the event (default). <strong>True</strong> indicates the application handles the event.                                                      |
| <em>ctrl</em>         | Required. The control being dragged over.                                                                                                                                                                                             |
| <em>Data</em>         | Required. Data that is dragged in a drag-and-drop operation. The data is packaged in a  <strong>DataObject</strong>.                                                                                                                  |
| <em>X, Y</em>         | Required. The horizontal and vertical coordinates of the control's position. Both coordinates are measured in points.  <em>X</em> is measured from the left edge of the control; <em>Y</em> is measured from the top of the control.. |
| <em>DragState</em>    | Required. Transition state of the data being dragged.                                                                                                                                                                                 |
| <em>Effect</em>       | Required. Operations supported by the [drop source](glossary-vba.md).                                                                                                                                                                 |
| <em>Shift</em>        | Required. Specifies the state of SHIFT, CTRL, and ALT.                                                                                                                                                                                |

 **Settings**
The settings for  _DragState_ are:


| <strong>Constant</strong> | <strong>Value</strong> | <strong>Description</strong>                                                     |
|:--------------------------|:-----------------------|:---------------------------------------------------------------------------------|
| <em>fmDragStateEnter</em> | 0                      | Mouse pointer is within range of a target.                                       |
| <em>fmDragStateLeave</em> | 1                      | Mouse pointer is outside the range of a target.                                  |
| <em>fmDragStateOver</em>  | 2                      | Mouse pointer is at a new position, but remains within range of the same target. |

The settings for  _Effect_ are:


| <strong>Constant</strong>       | <strong>Value</strong> | <strong>Description</strong>                              |
|:--------------------------------|:-----------------------|:----------------------------------------------------------|
| <em>fmDropEffectNone</em>       | 0                      | Does not copy or move the drop source to the drop target. |
| <em>fmDropEffectCopy</em>       | 1                      | Copies the drop source to the drop target.                |
| <em>fmDropEffectMove</em>       | 2                      | Moves the drop source to the drop target.                 |
| <em>fmDropEffectCopyOrMove</em> | 3                      | Copies or moves the drop source to the drop target.       |

The settings for  _Shift_ are:


| <strong>Constant</strong> | <strong>Value</strong> | <strong>Description</strong> |
|:--------------------------|:-----------------------|:-----------------------------|
| <em>fmShiftMask</em>      | 1                      | SHIFT was pressed.           |
| <em>fmCtrlMask</em>       | 2                      | CTRL was pressed.            |
| <em>fmAltMask</em>        | 4                      | ALT was pressed.             |

 **Remarks**
Use this event to monitor the mouse pointer as it enters, leaves, or rests directly over a valid [target](glossary-vba.md). When a drag-and-drop operation is in progress, the system initiates this event when the user moves the mouse, or presses or releases the mouse button or buttons. The mouse pointer position determines the target object that receives this event. You can determine the state of the mouse pointer by examining the  _DragState_ argument.
When a control handles this event, you can use the  _Effect_ argument to identify the drag-and-drop action to perform. When _Effect_ is set to **fmDropEffectCopyOrMove**, the drop source supports a copy ( **fmDropEffectCopy** ), move ( **fmDropEffectMove** ), or a cancel ( **fmDropEffectNone** ) operation.
When  _Effect_ is set to **fmDropEffectCopy**, the drop source supports a copy or a cancel ( **fmDropEffectNone** ) operation.
When  _Effect_ is set to **fmDropEffectMove**, the drop source supports a move or a cancel ( **fmDropEffectNone** ) operation.
When  _Effect_ is set to **fmDropEffectNone**. the drop source supports a cancel operation.
Most controls do not support drag-and-drop while  _Cancel_ is **False**, which is the default setting. This means the control rejects attempts to drag or drop anything on the control, and the control does not initiate the BeforeDropOrPaste event. The **TextBox** and **ComboBox** controls are exceptions to this; these controls support drag-and-drop operations even when _Cancel_ is **False**.

