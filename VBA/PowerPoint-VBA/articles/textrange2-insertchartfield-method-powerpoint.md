---
title: TextRange2.InsertChartField Method (PowerPoint)
ms.assetid: 42c07916-74e1-46c2-8cbc-5777c9fe1ae4
ms.date: 06/08/2017
ms.prod: powerpoint
---


# TextRange2.InsertChartField Method (PowerPoint)

Inserts a field into the body of a data label in a chart. 

This method applies only to data labels in a chart. Calling this method on any other kind of [TextRange2](http://msdn.microsoft.com/library/a6a59c9b-9b64-c1e2-2e98-a1f99025c877%28Office.15%29.aspx) object will raise a runtime error.

## Syntax

 _expression_. **InsertChartField**_(ChartFieldType,_ _Formula,_ _Position)_

 _expression_ A variable that represents a **TextRange2** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ChartFieldType_|Required|[MsoChartFieldType](http://msdn.microsoft.com/library/ce6b367d-d09f-4345-33e3-f181b1a9a41d%28Office.15%29.aspx)|Specifies the type of chart field to insert into a data label.|
| _Formula_|Optional|**string**|Specifies a cell (or range) if the  **MsoChartFieldFormula** constant is passed in for the _ChartFieldType_ parameter.|
| _Position_|Optional|**integer**|Specifies the character position where the chart field is inserted. The default is to append the field to the end of the text. If the position value is out of range, the default is used.|
| _ChartFieldType_|Required|MSOCHARTFIELDTYPE||
| _Formula_|Optional|STRING||
| _Position_|Optional|INT||
|Name|Required/Optional|Data type|Description|

### Return value

[TextRange2](http://msdn.microsoft.com/library/a6a59c9b-9b64-c1e2-2e98-a1f99025c877%28Office.15%29.aspx)


## See also


#### Concepts


[TextRange2 Object (PowerPoint)](textrange2-object-powerpoint.md)


