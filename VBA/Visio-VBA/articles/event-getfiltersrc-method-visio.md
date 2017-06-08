---
title: Event.GetFilterSRC Method (Visio)
keywords: vis_sdr.chm12650615
f1_keywords:
- vis_sdr.chm12650615
ms.prod: visio
api_name:
- Visio.Event.GetFilterSRC
ms.assetid: fcf9a5c1-cee9-df26-d774-df45c113945a
ms.date: 06/08/2017
---


# Event.GetFilterSRC Method (Visio)

Returns an array of cell ranges and a  **True** or **False** value indicating whether you are filtering events for that range.


## Syntax

 _expression_ . **GetFilterSRC**

 _expression_ A variable that represents an **Event** object.


### Return Value

Integer()


## Remarks

The event filters described in the array returned by the  **GetFilterSRC** method provide developers a way of ignoring specified events based on object type. The array returned is that passed to the **SetFilterSRC** method for this **Event** object.

The array that is returned by the  **GetFilterSRC** method can be interpreted in the following manner.

The number of elements in the array is a multiple of 7. These seven elements contain the following values:




- The first three elements describe the section, row, and cell of the beginning cell of the range.
    
- The next three elements describe the section, row, and cell of the end cell of the range.
    
- The last element contains a  **True** or **False** value indicating whether you want to receive events for the specified range of cells ( **True** to listen to events for a range of cells; **False** to exclude events for the range of cells).
    


For an event to successfully pass through a cell range filter, it must satisfy the following criteria:




- It must be a valid section, row, cell reference.
    
- If all filters are  **True** , the event must match at least one filter.
    
- If all filters are  **False** , the event must not match any filter.
    
- If the filters are a mixture of  **True** and **False** , the event must match at least one **True** filter and not match any **False** filters.
    


If there are no  **True** ranges defined in the array, events are considered **True** .

For details about using command IDs to define event filters, see the  **SetFilterSRC** method.


