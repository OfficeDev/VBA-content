
# Columns.Item Method (Publisher)

 **Last modified:** July 28, 2015

Returns an individual  **Column** object in the specified **Columns** collection.

## Syntax

 _expression_. **Item**( **_Index_**)

 _expression_A variable that represents a  **Columns** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Required| **Long**|The number of the object to return.|

### Return Value

Column


## Example

This example returns the first column from a  **Columns** collection.


```
Dim colTemp As Column 
 
Set colTemp = ActiveDocument.Pages(Index:=1) _ 
 .Shapes(1).Table.Columns.Item(Index:=1)
```

