
# TextRange2.InsertChartField Method (Office)

 **Last modified:** July 28, 2015

Inserts a field into the body of a data label in a chart. 

This method applies only to data labels in a chart. Calling this method on any other kind of  [TextRange2](a6a59c9b-9b64-c1e2-2e98-a1f99025c877.md) object will raise a runtime error.


## Syntax

 _expression_. **InsertChartField**(ChartFieldType,Formula,Position)

 _expression_A variable that represents a  **TextRange2** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
|ChartFieldType|Required| [MsoChartFieldType](ce6b367d-d09f-4345-33e3-f181b1a9a41d.md)|Specifies the type of chart field to insert into a data label.|
|Formula|Optional| **string**|Specifies a cell (or range) if the  **MsoChartFieldFormula** constant is passed in for theChartFieldType parameter.|
|Position|Optional| **integer**|Specifies the character position where the chart field is inserted. The default is to append the field to the end of the text. If the position value is out of range, the default is used.|
|ChartFieldType|Required|MSOCHARTFIELDTYPE||
|Formula|Optional|STRING||
|Position|Optional|INT||

### Return value

 [TextRange2](a6a59c9b-9b64-c1e2-2e98-a1f99025c877.md)

