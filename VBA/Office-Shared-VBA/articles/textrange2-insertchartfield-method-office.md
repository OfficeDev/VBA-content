---
title: TextRange2.InsertChartField Method (Office)
ms.assetid: 3ced5d2c-b3a4-6bf3-3d3c-b1145e7b9eab
ms.date: 06/08/2017
ms.prod: office
---


# TextRange2.InsertChartField Method (Office)

Inserts a field into the body of a data label in a chart. 

This method applies only to data labels in a chart. Calling this method on any other kind of [TextRange2](textrange2-object-office.md) object will raise a runtime error.

## Syntax

 _expression_. **InsertChartField**_(ChartFieldType,_ _Formula,_ _Position)_

 _expression_ A variable that represents a **TextRange2** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ChartFieldType_|Required|[MsoChartFieldType](msochartfieldtype-enumeration-office.md)|Specifies the type of chart field to insert into a data label.|
| _Formula_|Optional|**string**|Specifies a cell (or range) if the  **MsoChartFieldFormula** constant is passed in for the _ChartFieldType_ parameter.|
| _Position_|Optional|**integer**|Specifies the character position where the chart field is inserted. The default is to append the field to the end of the text. If the position value is out of range, the default is used.|
| _ChartFieldType_|Required|MSOCHARTFIELDTYPE||
| _Formula_|Optional|STRING||
| _Position_|Optional|INT||
|Name|Required/Optional|Data type|Description|

### Return value

[TextRange2](textrange2-object-office.md)


