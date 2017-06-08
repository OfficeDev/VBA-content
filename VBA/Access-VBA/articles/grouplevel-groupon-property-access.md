---
title: GroupLevel.GroupOn Property (Access)
keywords: vbaac10.chm12243
f1_keywords:
- vbaac10.chm12243
ms.prod: access
api_name:
- Access.GroupLevel.GroupOn
ms.assetid: 7fb9551f-5742-39a2-1cf3-7b3975ae517a
ms.date: 06/08/2017
---


# GroupLevel.GroupOn Property (Access)

You can use the  **GroupOn** property in a report to specify how to group data in a field or expression by data type. For example, this property lets you group a Date field by month. Read/write **Integer**.


## Syntax

 _expression_. **GroupOn**

 _expression_ A variable that represents a **GroupLevel** object.


## Remarks

The  **GroupOn** property settings available for a field depend on its data type, as the following table shows. For an expression, all of the settings are available. The default setting for all data types is Each Value.



|**Field data type**|**Setting**|**Groups records with**|**Visual Basic**|
|:-----|:-----|:-----|:-----|
|Text|(Default) Each Value|The same value in the field or expression.|0|
||Prefix Characters|The same first  _n_ number of characters in the field or expression.|1|
|Date/Time|(Default) Each Value|The same value in the field or expression.|0|
||Year|Dates in the same calendar year.|2|
||Qtr|Dates in the same calendar quarter.|3|
||Month|Dates in the same month.|4|
||Week|Dates in the same week.|5|
||Day|Dates on the same date.|6|
||Hour|Times in the same hour.|7|
||Minute|Times in the same minute.|8|
|AutoNumber, Currency, Number|(Default) Each Value|The same value in the field or expression.|0|
||Interval|Values within an interval you specify.|9|
In Visual Basic, you set this property in the  **[Open](report-open-event-access.md)** event procedure of a report.

To set the  **GroupOn** property to a value other than Each Value, you must first set the **GroupHeader** or **GroupFooter** property or both to Yes for the selected field or expression.


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

