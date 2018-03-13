---
title: Event.SetFilterActions Method (Visio)
keywords: vis_sdr.chm12660260
f1_keywords:
- vis_sdr.chm12660260
ms.prod: visio
api_name:
- Visio.Event.SetFilterActions
ms.assetid: 8a0f7b5c-466b-7b98-a34f-6a639fded39c
ms.date: 06/08/2017
---


# Event.SetFilterActions Method (Visio)

Specifies the extensions to the  **MouseMove** event that Visio reports.


## Syntax

 _expression_ . **SetFilterActions**( **_filterActionStream()_** )

 _expression_ An expression that returns a **Event** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _filterActionStream()_|Required| **Long**|An array of action/value pairs. For more information, see Remarks.|

### Return Value

Nothing


## Remarks

The  **SetFilterActions** method provides a way of ignoring selected extensions of the **MouseMove** event based on extension type. Extension types are based on mouse actions that are part of a drag and drop operation, as shown in the table below. By default, Visio reports firings of all event extensions.

The  _filterActionStream_ parameter is an array defined in the following way. The number of elements in _filterActionStream_ is a multiple of 3:


- The first element contains the beginning mouse action ( **MouseMove** event extension) of the range (any member of **VisFilterActions** ).

- The second element contains the end mouse action ( **MouseMove** event extension) of the range (any member of **VisFilterActions** whose value is higher than that of the first element ).

- The third element contains a  **True** or **False** value indicating whether you want to listen to events for that action range ( **True** to listen to events of a certain sub-type, or **MouseMove** event extension; **False** to exclude an event sub-type).

The filter actions that you can place in the first and second array elements of each element triplet are defined in the  **VisFilterActions** enumeration, which is declared in the Visio type library, and shown in the following table.



| <strong>Constant</strong>                    | <strong>Value</strong> | <strong>Description</strong>                                                              |
|:---------------------------------------------|:-----------------------|:------------------------------------------------------------------------------------------|
| <strong>visFilterMouseMoveDragBegin</strong> | 1                      | Filter the  <strong>DragBegin</strong> extension of the <strong>MouseMove</strong> event. |
| <strong>visFilterMouseMoveDragDrop</strong>  | 5                      | Filter the  <strong>DragDrop</strong> extension of the <strong>MouseMove</strong> event.  |
| <strong>visFilterMouseMoveDragEnter</strong> | 2                      | Filter the  <strong>DragEnter</strong> extension of the <strong>MouseMove</strong> event. |
| <strong>visFilterMouseMoveDragLeave</strong> | 4                      | Filter the  <strong>DragLeave</strong> extension of the <strong>MouseMove</strong> event. |
| <strong>visFilterMouseMoveDragOver</strong>  | 3                      | Filter the  <strong>DragOver</strong> extension of the <strong>MouseMove</strong> event.  |
| <strong>visFilterMouseMoveNoDrag</strong>    | 0                      | Do not filter any extensions of the  <strong>MouseMove</strong> event.                    |

For example, if you want to listen to all  **MouseEvent** extensions except the **DragOver** event extension, you can build an array like the following:




```vb
Dim alngFilterActions(1 to 1 * 3) As Long  
    alngFilterActions(1) = visFilterMouseMoveDragDrop  
    alngFilterActions(2) = visFilterMouseMoveDragDrop  
    alngFilterActions(3) = False 
```

Or, to listen only to the  **DragEnter** event extension, ignoring mouse actions that come before and after, set up an array like the following:




```vb
Dim alngFilterActions(1 To (3 * 3)) As Long  

    'Listen to the "DragEnter" mouse action.  
    alngFilterActions(1) = visFilterMouseMoveDragEnter  
    alngFilterActions(2) = visFilterMouseMoveDragEnter   
    alngFilterActions(3) = True  

    'Ignore any mouse actions before "DragEnter."   
    alngFilterActions(4) = visFilterMouseMoveDragBegin  
    alngFilterActions(5) = visFilterMouseMoveDragEnter  - 1  
    alngFilterActions(6) = False  

    'Ignore any mouse actions after "DragEnter."   
    alngFilterActions(7) = visFilterMouseMoveDragEnter + 1  
    alngFilterActions(8) = visFilterMouseMoveDragDrop  
    alngFilterActions(9) = False 
```

Note that mouse actions that occupy the second position in an array-element triplet must always be later in the sequence (that is, higher in value) than those that occupy the first position in an array-element triplet.


