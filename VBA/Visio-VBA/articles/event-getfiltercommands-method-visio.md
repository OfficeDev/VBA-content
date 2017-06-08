---
title: Event.GetFilterCommands Method (Visio)
keywords: vis_sdr.chm12650610
f1_keywords:
- vis_sdr.chm12650610
ms.prod: visio
api_name:
- Visio.Event.GetFilterCommands
ms.assetid: 47664b2f-702b-1c61-1746-9b5fd470a8f4
ms.date: 06/08/2017
---


# Event.GetFilterCommands Method (Visio)

Returns an array of command ranges and a  **True** or **False** value indicating how to filter events for that command range.


## Syntax

 _expression_ . **GetFilterCommands**

 _expression_ A variable that represents an **Event** object.


### Return Value

Long()


## Remarks

The event filters described in the array returned by the  **GetFilterCommands** method provide developers a way of ignoring specified events based on command ID. The array returned is that passed to the **SetFilterCommands** method for this **Event** object.

The array that is returned by the  **GetFilterCommands** method can be interpreted in the following manner:

The number of elements in the array is a multiple of 3, as follows:




- The first element contains the beginning command ID of the range (any member of  **[VisUICmds](visuicmds-enumeration-visio.md)** ).
    
- The second element contains the end command ID of the range (any member of  **VisUICmds** ).
    
- The third element contains a  **True** or **False** value, which indicates whether you are listening to events for that command range ( **True** to listen to events; **False** to exclude events).
    


For an event to successfully pass through a command filter, it must satisfy the following criteria:




- It must have a valid command ID.
    
- If all filters are  **True** , the event must match at least one filter.
    
- If all filters are  **False** , the event must not match any filter.
    
- If the filters are a mixture of  **True** and **False** , the event must match at least one **True** filter and not match any **False** filters.
    


If there are no  **True** ranges defined in the array, events are considered **True** .

For details about using command IDs to define event filters, see the  **SetFilterCommands** method.


