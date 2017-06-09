---
title: Event.GetFilterObjects Method (Visio)
keywords: vis_sdr.chm12650605
f1_keywords:
- vis_sdr.chm12650605
ms.prod: visio
api_name:
- Visio.Event.GetFilterObjects
ms.assetid: 89df76cd-16a4-d3ba-7399-b4dfca49b9cb
ms.date: 06/08/2017
---


# Event.GetFilterObjects Method (Visio)

Returns an array of object types and a  **True** or **False** value indicating how to filter events for that object.


## Syntax

 _expression_ . **GetFilterObjects**

 _expression_ A variable that represents an **Event** object.


### Return Value

Long()


## Remarks

The event filters described in the array returned by the  **GetFilterObjects** method provide developers a way of ignoring specified events based on object type. The array returned is that passed to the **SetFilterObjects** method for this **Event** object.

The array that is returned by the  **GetFilterObjects** method can be interpreted in the following manner.

The number of elements in the array is a multiple of 2: 


- The first element contains an object type (one of  **visTypePage** , **visTypeGroup** , **visTypeShape** , **visTypeForeignObject** , **visTypeGuide** , or **visTypeDoc** ).
    
- The second element contains a  **True** or **False** value indicating whether you are listening to events for that object ( **True** to listen to an object's events; **False** to exclude an object's events).
    


For an event to successfully pass through an object event filter, it must satisfy the following criteria: 


- It must be a valid object type.
    
- If all filters are  **True** , the event must match at least one filter.
    
- If all filters are  **False** , the event must not match any filter.
    
- If the filters are a mixture of  **True** and **False** , the event must match at least one **True** filter and not match any **False** filters.
    


If there are no  **True** ranges defined in the array, events are considered **True** .

For details about defining event filters using command IDs, see the  **SetFilterObjects** method.


