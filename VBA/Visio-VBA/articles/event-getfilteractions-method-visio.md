---
title: Event.GetFilterActions Method (Visio)
keywords: vis_sdr.chm12660255
f1_keywords:
- vis_sdr.chm12660255
ms.prod: visio
api_name:
- Visio.Event.GetFilterActions
ms.assetid: c74be758-280a-13a8-5462-b508bd3f50e4
ms.date: 06/08/2017
---


# Event.GetFilterActions Method (Visio)

Returns an array of the filter actions set for the  **Event** object.


## Syntax

 _expression_ . **GetFilterActions**

 _expression_ An expression that returns a **Event** object.


### Return Value

Long()


## Remarks

The event filters described in the array returned by the  **GetFilterActions** method provide developers a way of ignoring specified mouse-event extensions based on extension (action) type. The array returned is that passed to the **SetFilterActions** method for this **Event** object. The array that is returned by the **GetFilterActions** method can be interpreted in the following manner.

The number of elements in the array is a multiple of 3, as follows:


- The first element contains the beginning mouse action ( **MouseMove** event extension) of the range (any member of **VisFilterActions** ).
    
- The second element contains the end mouse action ( **MouseMove** event extension) of the range (any member of **VisFilterActions** whose value is higher than that of the first element ).
    
- The third element contains a  **True** or **False** value indicating whether you want to listen to events for that action range ( **True** to listen to events of a certain sub-type, or **MouseMove** event extension; **False** to exclude an event sub-type).
    
The filter actions that are returned in the first and second array elements of each element triplet are defined in the  **VisFilterActions** enumeration, which is declared in the Visio type library, and shown in the following table. Note that mouse actions that occupy the second position in an array-element triplet will always be later in the sequence (that is, higher in value) than those that occupy the first position in an array-element triplet.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visFilterMouseMoveDragBegin**|1|Filter the  **DragBegin** extension of the **MouseMove** event.|
| **visFilterMouseMoveDragDrop**|5|Filter the  **DragDrop** extension of the **MouseMove** event.|
| **visFilterMouseMoveDragEnter**|2|Filter the  **DragEnter** extension of the **MouseMove** event.|
| **visFilterMouseMoveDragLeave**|4|Filter the  **DragLeave** extension of the **MouseMove** event.|
| **visFilterMouseMoveDragOver**|3|Filter the  **DragOver** extension of the **MouseMove** event.|
| **visFilterMouseMoveNoDrag**|0|Do not filter any extensions of the  **MouseMove** event.|
For more information about using event extensions to define filter actions, see the  **[SetFilterActions](event-setfilteractions-method-visio.md)** method.


