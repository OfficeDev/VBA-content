---
title: Reports.Item Property (Project)
ms.prod: project-server
ms.assetid: d8202579-71de-c606-5a28-af285bca0a05
ms.date: 06/08/2017
---


# Reports.Item Property (Project)
Gets a single  **Report** object from the **Reports** collection. Read-only **Report**.

## Syntax

 _expression_. **Item**

 _expression_ A variable that represents a **Reports** object.


## Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Variant**|Name of the report or index number of the report.|

## Remarks

To get the index number of a report, you can use the [Report.Index](report-index-property-project.md) property. For example, create a report namedReport 1, and then run the following statement in the  **Immediate** window of the VBE:


```vb
? ActiveProject.Reports.Item("Report 1").Index
```

 **Item** is the default property of the **Reports** object. For example, the following statement is equivalent of the previous statement.




```vb
? ActiveProject.Reports("Report 1").Index
```


## Property value

 **REPORT**


## See also


#### Other resources


[Reports Object](reports-object-project.md)
