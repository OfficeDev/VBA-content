---
title: Report.Requery Method (Access)
keywords: vbaac10.chm13827
f1_keywords:
- vbaac10.chm13827
ms.prod: ACCESS
api_name:
- Access.Report.Requery
ms.assetid: d078d523-3bbd-fa70-44ac-116cdcedfebd
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


 **Note**  


## See also


#### Concepts


[Report Object](report-object-access.md)

