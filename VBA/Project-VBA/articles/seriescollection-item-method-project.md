---
title: SeriesCollection.Item Method (Project)
ms.prod: project-server
ms.assetid: 3360bb21-9494-f39d-91e8-049a8fae6ad5
ms.date: 06/08/2017
---


# SeriesCollection.Item Method (Project)
Gets an individual  **Series** object in the series collection. Read-only **Series**.

## Syntax

 _expression_. **Item** _(Index)_

 _expression_ A variable that represents a **SeriesCollection** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Variant**|The number or name of the series collection.|

### Example

The following example prints the name of the first series in the series collection of the specified active report, to the  **Immediate** window of the VBE.


```vb
? ActiveProject.Reports("Simple scalar chart").Shapes(1).Chart.SeriesCollection.Item(1).Name
```

The  **Item** method is not required in some cases; for example, the following example has the same result.




```vb
? ActiveProject.Reports("Simple scalar chart").Shapes(1).Chart.SeriesCollection(1).Name
```


