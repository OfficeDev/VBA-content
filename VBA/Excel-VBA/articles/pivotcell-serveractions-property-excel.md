---
title: PivotCell.ServerActions Property (Excel)
keywords: vbaxl10.chm692090
f1_keywords:
- vbaxl10.chm692090
ms.prod: excel
ms.assetid: e895f7ee-e636-29b6-9385-2710885cc01c
ms.date: 06/08/2017
---


# PivotCell.ServerActions Property (Excel)

Represents a collection of  _actions_ consisting of OLAP-defined actions which can be executed. The actions are specific to PivotTables existing at a worksheet-level. Read-only


## Syntax

 _expression_ . **ServerActions**

 _expression_ A variable that represents a **PivotCell** object.


## Remarks

A server action is an optional feature that an OLAP cube administrator can define on a server that uses a cube member or measure as a parameter into a query to obtain details in the cube.


## Example

The following code segment executes a server action against a series in a PivotChart.


```vb
ActiveSheet.ChartObjects("Chart 1").Chart.PivotLayout.PivotTable.PivotColumnAxis.PivotLines(index of line ).PivotLineCells(index of cells ).ServerAction("OLAP Action name" ).Execute
```


## Property value

 **ACTIONS**


## See also


#### Concepts


[PivotCell Object](pivotcell-object-excel.md)

