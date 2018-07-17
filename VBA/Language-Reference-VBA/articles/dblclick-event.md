---
title: DblClick Event
keywords: fm20.chm5224940
f1_keywords:
- fm20.chm5224940
ms.prod: office
api_name:
- Office.DblClick
ms.assetid: 52ee3887-6634-ed57-fb9b-757543ea6e29
ms.date: 06/08/2017
---


# DblClick Event



Occurs when the user points to an object and then clicks a mouse button twice.
 **Syntax**
For MultiPage, TabStrip **Private Sub**_object_ _**DblClick(**_index_**As Long**, **ByVal**_Cancel_**As MSForms.ReturnBoolean)**
For other controls **Private Sub**_object_ _**DblClick( ByVal**_Cancel_**As MSForms.ReturnBoolean)**
The  **DblClick** event syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _index_|Required. The position of a  **Page** or **Tab** object within a **Pages** or **Tabs** collection.|
| _Cancel_|Required. Event status.  **False** indicates that the control should handle the event (default). **True** indicates the application handles the event.|
 **Remarks**
For this event to occur, the two clicks must occur within the time span specified by the system's double-click speed setting.
For controls that support Click, the following sequence of events leads to the DblClick event:


1. MouseDown
    
2. MouseUp
    
3. Click
    
4. DblClick
    

If a control, such as  **TextBox**, does not support Click, Click is omitted fom the order of events leading to the DblClick event.
If the return value of  _Cancel_ is **True** when the user clicks twice, the control ignores the second click. This is useful if the second click reverses the effect of the first, such as double-clicking a toggle button. The _Cancel_ argument allows your form to ignore the second click, so that either clicking or double-clicking the button has the same effect.

