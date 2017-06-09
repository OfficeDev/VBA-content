---
title: Application.SelectTable Method (Project)
keywords: vbapj.chm1516
f1_keywords:
- vbapj.chm1516
ms.prod: project-server
ms.assetid: 8cf26b2d-4021-cf2a-8f0d-d033965f3629
ms.date: 06/08/2017
---


# Application.SelectTable Method (Project)
Selects the entire table, when one or more items within a table shape are selected in a report.

## Syntax

 _expression_. **SelectTable**

 _expression_ A variable that represents an **Application** object.


### Return value

 **Boolean**

 **True** if the table is selected; otherwise, **False**.


## Remarks

When one or more items within a table shape are selected, the  **SelectTable** method removes the selection highlighting of the item(s), and selects the entire table.

If the active view is not a report, the  **SelectTable** method displays a dialog box with run-time error 1100, "The method is not available in this situation." If no item in a report is selected, or if an item in another kind of shape (not a table) is selected, Project shows run-time error -2147417848, "Automation error. The object invoked has disconnected from its clients."


## See also


#### Concepts


[Application Object](application-object-project.md)
#### Other resources


[ReportTable Object](reporttable-object-project.md)
[Shape.Table Property](shape-table-property-project.md)
