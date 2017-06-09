---
title: Report.Filter Event (Access)
keywords: vbaac10.chm13898
f1_keywords:
- vbaac10.chm13898
ms.prod: access
api_name:
- Access.Report.Filter
ms.assetid: 1344ceff-d3ac-3dc1-0f9c-563d895a77dc
ms.date: 06/08/2017
---


# Report.Filter Event (Access)

Occurs when the user opens a filter window by clicking  **Advanced Filter/Sort**.


## Syntax

 _expression_. **Filter**( ** _Cancel_**, ** _FilterType_** )

 _expression_ A variable that represents a **Report** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required|**Integer**|The setting determines whether to open the filter window. Setting the  _Cancel_ argument to **True** (?1) prevents the filter window from opening. You can also use the **CancelEvent** method of the **DoCmd** object to cancel opening the filter window.|
| _FilterType_|Required|**Integer**|The filter window the user is trying to open. The  _FilterType_ argument can be **acFilterAdvanced**.|

## Remarks

To run a macro or event procedure when this event occurs, set the  **OnFilter** property to the name of the macro or to [Event Procedure].

You can use the  **Filter** event to:


- Remove any previous filter for the report. To do this, set the  **Filter** property or **ServerFilter** property of the report to a zero-length string (" ") in the **Filter** macro or event procedure. This is especially useful if you want to make sure extraneous criteria do not appear in the new filter. For example, when you use the Filter By Selection feature, the criteria you use (the selected text in the report) is added to the **Filter** or **ServerFilter** property WHERE clause expression, and appears in the **Advanced Filter/Sort** window. You can remove these old criteria by using the **Filter** event.
    
- Enter default settings for the new filter. To do this, set the  **Filter** property or **ServerFilter** property to include these criteria. For example, you may want all filters for a Products report to display only current products (products for which the Discontinued control in the Products report is not selected).
    
- Use your own custom filter window instead of one of the Microsoft Access filter windows. When the  **Filter** event occurs, you can open your own custom form and use the entries on this report to set the **Filter** property or **ServerFilter** property and filter the original report. When the user closes this custom form, set the **FilterOn** property or **ServerFilterByForm** property of the original report to **True** (?1) to apply the filter. Canceling the **Filter** event prevents the Microsoft Access filter window from opening.
    

## See also


#### Concepts


[Report Object](report-object-access.md)

