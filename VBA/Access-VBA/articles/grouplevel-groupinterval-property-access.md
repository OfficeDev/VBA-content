---
title: GroupLevel.GroupInterval Property (Access)
keywords: vbaac10.chm12244
f1_keywords:
- vbaac10.chm12244
ms.prod: access
api_name:
- Access.GroupLevel.GroupInterval
ms.assetid: 98ba66b9-658e-9fa6-155b-3b4f1a7c3f01
ms.date: 06/08/2017
---


# GroupLevel.GroupInterval Property (Access)

You can use the  **GroupInterval** property with the **[GroupOn](grouplevel-groupon-property-access.md)** property to specify how records are grouped in a report. Read/write **Long**.


## Syntax

 _expression_. **GroupInterval**

 _expression_ A variable that represents a **GroupLevel** object.


## Remarks

The  **GroupInterval** property specifies an interval value that records are grouped by. This interval differs depending on the data type and **GroupOn** property setting of the field or expression you're grouping on. For example, you can set the **GroupInterval** property to 1 if you want to group records by the first character of a Text field, such as ProductName.

The  **GroupInterval** property settings are **Long** values that depend on the field's data type and its **GroupOn** property setting. The default **GroupInterval** setting is 1.

You can set the  **GroupInterval** property only the **[Open](report-open-event-access.md)** event procedure of a report.


 **Note**  To set the  **GroupInterval** property to a value other than its default setting (1), you must first set the **[GroupHeader](grouplevel-groupheader-property-access.md)** or **[GroupFooter](grouplevel-groupfooter-property-access.md)** property or both to Yes for the selected field or expression.

Here are examples of  **GroupInterval** property settings for different field data types.



|**Field data type**|**GroupOn setting**|**GroupInterval setting**|
|:-----|:-----|:-----|
|All|Each value|(Default) Set to 1.|
|Text|Prefix characters|Set to 3 for grouping by the first three characters in the field (for example, Chai, Chartreuse, and Chang would be grouped together).|
|Date/Time|Week|Set to 2 to return data in biweekly groups.|
|Date/Time|Hour|Set to 12 to return data in half-day groups.|

## Example

The following example sets the  **SortOrder** and grouping properties for the first group level in the Products By Category report to create an alphabetical list of products.


```vb
Private Sub Report_Open(Cancel As Integer) 
 ' Set SortOrder property to ascending order. 
 Me.GroupLevel(0).SortOrder = False 
 ' Set GroupOn property. 
 Me.GroupLevel(0).GroupOn = 1 
 ' Set GroupInterval property to 1. 
 Me.GroupLevel(0).GroupInterval = 1 
 ' Set KeepTogether property to With First Detail. 
 Me.GroupLevel(0).KeepTogether = 2 
End Sub
```


## See also


#### Concepts


[GroupLevel Object](grouplevel-object-access.md)

