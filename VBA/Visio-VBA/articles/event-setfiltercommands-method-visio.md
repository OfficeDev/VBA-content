---
title: Event.SetFilterCommands Method (Visio)
keywords: vis_sdr.chm12650830
f1_keywords:
- vis_sdr.chm12650830
ms.prod: visio
api_name:
- Visio.Event.SetFilterCommands
ms.assetid: ef4550c8-e77b-032f-52b2-7e2d18b2316f
ms.date: 06/08/2017
---


# Event.SetFilterCommands Method (Visio)

Specifies an array of command ranges and a  **True** or **False** value indicating how to filter events for each command range.


## Syntax

 _expression_ . **SetFilterCommands**( **_Commands()_** )

 _expression_ A variable that represents an **Event** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Commands()_|Required| **Long**|An array of command ranges and a  **True** or **False** value specifying how to filter events for each command range.|

### Return Value

Nothing


## Remarks

When an  **Event** object created with the **AddAdvise** method is added to the **EventList** collection of a source object, the default behavior is that all occurrences of that event are passed to the event sink. The **SetFilterCommands** method provides a way of ignoring selected events based on command ID.

The  _Commands()_ parameter passed to **SetFilterCommands** is an array defined in the following way.

The number of elements in  _Commands()_ is a multiple of 3:


- The first element contains the beginning command ID of the range (any member of  **VisUICmds** ).
    
- The second element contains the end command ID of the range (any member of  **VisUICmds** ).
    
- The third element contains a  **True** or **False** value, which indicates whether you are listening to events for that command range ( **True** to listen to events; **False** to exclude events).
    


For an event to successfully pass through a command filter, it must satisfy the following criteria:




- It must have a valid command ID.
    
- If all filters are  **True** , the event must match at least one filter.
    
- If all filters are  **False** , the event must not match any filter.
    
- If the filters are a mixture of  **True** and **False** , the event must match at least one **True** filter and not match any **False** filters.
    


If there are no  **True** ranges in the array, events are considered **True** .

For example, to set up an array that blocks out a single command, use the following: 




```vb
 
    Dim aFilterCommands(1 To (1 * 3)) As Long  
 
    'Ignore the layout command. 
    aFilterCommands(1) = visCmdLayoutDynamic  
    aFilterCommands(2) = visCmdLayoutDynamic  
    aFilterCommands(3) = False 

```

Or, to set up an array that listens only to the  **Send to Back** command:




```vb
 
    Dim aFilterCommands(1 To (3 * 3)) As Long  
 
    'Pay attention to the "Send to Back" command.  
    aFilterCommands(1) = visCmdObjectSendToBack  
    aFilterCommands(2) = visCmdObjectSendToBack  
    aFilterCommands(3) = True  
 
    'Ignore any command IDs before the "Send to Back" command.  
    aFilterCommands(4) = visCmdFirst  
    aFilterCommands(5) = visCmdObjectSendToBack - 1  
    aFilterCommands(6) = False  
 
    'Ignore any command IDs after the "Send to Back" command.  
    aFilterCommands(7) = visCmdObjectSendToBack + 1  
    aFilterCommands(8) = visCmdLast  
    aFilterCommands(9) = False 

```


