---
title: Report.Requery Method (Access)
keywords: vbaac10.chm13827
f1_keywords:
- vbaac10.chm13827
ms.prod: access
api_name:
- Access.Report.Requery
ms.assetid: d078d523-3bbd-fa70-44ac-116cdcedfebd
ms.date: 06/08/2017
---


# Report.Requery Method (Access)

The  **Requery** method updates the data underlying the specified report by requerying the source of data for the control.


## Syntax

 _expression_. **Requery**

 _expression_ A variable that represents a **Report** object.


## Remarks

You can use this method to ensure that a form or control displays the most recent data.

The  **Requery** method does one of the following:


- Reruns the query on which the report is based.
    
- Updates records displayed based on any changes to the  **Filter** property of the report.
    
If you omit the object specified by expression, the  **Requery** method requeries the underlying data source for the report that has the focus. If the control that has the focus has a record source or row source, it will be requeried; otherwise, the control's data will simply be refreshed.


|**Note**|
|:-----|
|<ul><li>The **Requery** method updates the data underlying a form or control to reflect records that are new to or deleted from the record source since it was last queried. The **Refresh** method shows only changes that have been made to the current set of records; it doesn't reflect new or deleted records in the record source. The **Repaint** method simply repaints the specified form and its controls.</li><li>The **Requery** method doesn't pass control to the operating system to allow Windows to continue processing messages. Use the **DoEvents** function if you need to relinquish temporary control to the operating system.</li><li>The **Requery** method is faster than the **Requery** action. When you use the **Requery** action, Microsoft Access closes the query and reloads it from the database. When you use the **Requery** method, Microsoft Access reruns the query without closing and reloading it.</li></ul>| 


## See also


#### Concepts


[Report Object](report-object-access.md)

