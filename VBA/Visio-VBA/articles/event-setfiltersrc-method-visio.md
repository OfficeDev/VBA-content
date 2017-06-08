---
title: Event.SetFilterSRC Method (Visio)
keywords: vis_sdr.chm12650840
f1_keywords:
- vis_sdr.chm12650840
ms.prod: visio
api_name:
- Visio.Event.SetFilterSRC
ms.assetid: 06ba59d2-57a4-7686-3250-388e499bfc76
ms.date: 06/08/2017
---


# Event.SetFilterSRC Method (Visio)

Specifies an array of cell ranges and a  **True** or **False** value indicating how to filter events for each cell range.


## Syntax

 _expression_ . **SetFilterSRC**( **_SRCStream()_** )

 _expression_ A variable that represents an **Event** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SRCStream()_|Required| **Integer**|An array of cell ranges and a  **True** or **False** value specifying how to filter events for each range.|

### Return Value

Nothing


## Remarks

When an  **Event** object created with the **AddAdvise** method is added to the **EventList** collection of a source object, the default behavior is that all occurrences of that event are passed to the event sink. The **SetFilterSRC** method provides a way to ignore selected events based on a range of cells.

The  _SRCStream()_ parameter passed to **SetFilterSRC** is an array defined in the following manner:

The number of elements in the array is a multiple of 7:




- The first three elements describe the section, row, and cell of the beginning cell of the range.
    
- The next three elements describe the section, row, and cell of the end cell of the range.
    
- The last element contains a  **True** or **False** value indicating how to filter events for the cell range ( **True** to listen to events for a range of cells; **False** to exclude events for a range of cells).
    


For an event to successfully pass through a cell range filter, it must satisfy the following criteria:




- It must be a valid section, row, cell reference.
    
- If all filters are  **True** , the event must match at least one filter.
    
- If all filters are  **False** , the event must not match any filter.
    
- If the filters are a mixture of  **True** and **False** , the event must match at least one **True** filter and not match any **False** filters.
    


If there are no  **True** ranges defined in the array, events are considered **True** .

For example, if you want to listen for any changes in the Value cell of the second row in the Shape Data section, use the following:




```vb
 
 Dim aFilterSRC(1 To (1 * 7)) As Integer 
 aFilterSRC(1) = visSectionProp 
 aFilterSRC(2) = visRowProp + 1 
 aFilterSRC(3) = visCustPropsValue 
 aFilterSRC(4) = visSectionProp 
 aFilterSRC(5) = visRowProp + 1 
 aFilterSRC(6) = visCustPropsValue 
 aFilterSRC(7) = True 

```


