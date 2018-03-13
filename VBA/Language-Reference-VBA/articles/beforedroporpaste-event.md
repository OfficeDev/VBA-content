---
title: BeforeDropOrPaste Event
keywords: fm20.chm5224936
f1_keywords:
- fm20.chm5224936
ms.prod: office
api_name:
- Office.BeforeDropOrPaste
ms.assetid: ba572265-1a9d-2d02-6346-82f88c1f249a
ms.date: 06/08/2017
---


# BeforeDropOrPaste Event



Occurs when the user is about to drop or paste data onto an object.
 <strong>Syntax</strong>
For Frame 
<strong>Private Sub</strong><em>object</em> <em><strong>BeforeDropOrPaste( ByVal</strong>_Cancel</em><strong>As MSForms.ReturnBoolean</strong>, <em>ctrl</em><strong>As Control</strong>, <strong>ByVal</strong><em>Action</em><strong>As fmAction</strong>, <strong>ByVal</strong><em>Data</em><strong>As DataObject</strong>, <strong>ByVal</strong><em>X</em><strong>As Single</strong>, <strong>ByVal</strong><em>Y</em><strong>As Single</strong>, <strong>ByVal</strong><em>Effect</em><strong>As MSForms.ReturnEffect</strong>, <strong>ByVal</strong><em>Shift</em><strong>As fmShiftState)</strong>
For MultiPage 
<strong>Private Sub</strong><em>object</em> <em><strong>BeforeDropOrPaste(</strong>_index</em><strong>As Long</strong>, <strong>ByVal</strong><em>Cancel</em><strong>As MSForms.ReturnBoolean</strong>, <em>ctrl</em><strong>As Control</strong>, <strong>ByVal</strong><em>Action</em><strong>As fmAction</strong>, <strong>ByVal</strong><em>Data</em><strong>As DataObject</strong>, <strong>ByVal</strong><em>X</em><strong>As Single</strong>, <strong>ByVal</strong><em>Y</em><strong>As Single</strong>, <strong>ByVal</strong><em>Effect</em><strong>As MSForms.ReturnEffect</strong>, <strong>ByVal</strong><em>Shift</em><strong>As fmShiftState)</strong>
For TabStrip 
<strong>Private Sub</strong><em>object</em> <em><strong>BeforeDropOrPaste(</strong>_index</em><strong>As Long</strong>, <strong>ByVal</strong><em>Cancel</em><strong>As MSForms.ReturnBoolean</strong>, <strong>ByVal</strong><em>Action</em><strong>As fmAction</strong>, <strong>ByVal</strong><em>Data</em><strong>As DataObject</strong>, <strong>ByVal</strong><em>X</em><strong>As Single</strong>, <strong>ByVal</strong><em>Y</em><strong>As Single</strong>, <strong>ByVal</strong><em>Effect</em><strong>As MSForms.ReturnEffect</strong>, <strong>ByVal</strong><em>Shift</em><strong>As fmShiftState)</strong>
For other controls 
<strong>Private Sub</strong><em>object</em> <em><strong>BeforeDropOrPaste( ByVal</strong>_Cancel</em><strong>As MSForms.ReturnBoolean</strong>, <strong>ByVal</strong><em>Action</em><strong>As fmAction</strong>, <strong>ByVal</strong><em>Data</em><strong>As DataObject</strong>, <strong>ByVal</strong><em>X</em><strong>As Single</strong>, <strong>ByVal</strong><em>Y</em><strong>As Single</strong>, <strong>ByVal</strong><em>Effect</em><strong>As MSForms.ReturnEffect</strong>, <strong>ByVal</strong><em>Shift</em><strong>As fmShiftState)</strong>
The  
<strong>BeforeDropOrPaste</strong> event syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong>                                                                                                                                                                                                                       |
|:----------------------|:---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>object</em>       | Required. A valid object name.                                                                                                                                                                                                                     |
| <em>index</em>        | Required. The index of the  <strong>Page</strong> in a <strong>MultiPage</strong> that the drop or paste operation will affect.                                                                                                                    |
| <em>Cancel</em>       | Required. Event status.  <strong>False</strong> indicates that the control should handle the event (default). <strong>True</strong> indicates the application handles the event.                                                                   |
| <em>ctrl</em>         | Required. The target control.                                                                                                                                                                                                                      |
| <em>Action</em>       | Required. Indicates the result, based on the current keyboard settings, of the pending drag-and-drop operation.                                                                                                                                    |
| <em>Data</em>         | Required. Data that is dragged in a drag-and-drop operation. The data is packaged in a  <strong>DataObject</strong>.                                                                                                                               |
| <em>X, Y</em>         | Required. The horizontal and vertical position of the mouse pointer when the drop occurs. Both coordinates are measured in points.  <em>X</em> is measured from the left edge of the control; <em>Y</em> is measured from the top of the control.. |
| <em>Effect</em>       | Required. Effect of the drag-and-drop operation on the target control.                                                                                                                                                                             |
| <em>Shift</em>        | Required. Specifies the state of SHIFT, CTRL, and ALT.                                                                                                                                                                                             |

 **Settings**
The settings for  _Action_ are:


| <strong>Constant</strong> | <strong>Value</strong> | <strong>Description</strong>                                                                                    |
|:--------------------------|:-----------------------|:----------------------------------------------------------------------------------------------------------------|
| <em>fmActionPaste</em>    | 2                      | Pastes the selected object into the drop target.                                                                |
| <em>fmActionDragDrop</em> | 3                      | Indicates the user has dragged the object from its source to the drop target and dropped it on the drop target. |

The settings for  _Effect_ are:


| <strong>Constant</strong>       | <strong>Value</strong> | <strong>Description</strong>                                                 |
|:--------------------------------|:-----------------------|:-----------------------------------------------------------------------------|
| <em>fmDropEffectNone</em>       | 0                      | Does not copy or move the [drop source](glossary-vba.md) to the drop target. |
| <em>fmDropEffectCopy</em>       | 1                      | Copies the drop source to the drop target.                                   |
| <em>fmDropEffectMove</em>       | 2                      | Moves the drop source to the drop target.                                    |
| <em>fmDropEffectCopyOrMove</em> | 3                      | Copies or moves the drop source to the drop target.                          |

The settings for  _Shift_ are:


| <strong>Constant</strong> | <strong>Value</strong> | <strong>Description</strong> |
|:--------------------------|:-----------------------|:-----------------------------|
| <em>fmShiftMask</em>      | 1                      | SHIFT was pressed.           |
| <em>fmCtrlMask</em>       | 2                      | CTRL was pressed.            |
| <em>fmAltMask</em>        | 4                      | ALT was pressed.             |

 **Remarks**
For a  **MultiPage** or **TabStrip**, Visual Basic for Applications initiates this event when it transfers a data object to the control.
For other controls, the system initiates this event immediately prior to the drop or paste operation.
When a control handles this event, you can update the  _Action_ argument to identify the drag-and-drop action to perform. When _Effect_ is set to **fmDropEffectCopyOrMove**, you can assign _Action_ to **fmDropEffectNone**, **fmDropEffectCopy**, or **fmDropEffectMove**. When _Effect_ is set to **fmDropEffectCopy** or **fmDropEffectMove**, you can reassign _Action_ to **fmDropEffectNone**. You cannot reassign _Action_ when _Effect_ is set to **fmDropEffectNone**.

