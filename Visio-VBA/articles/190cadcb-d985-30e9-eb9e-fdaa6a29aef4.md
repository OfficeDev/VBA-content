
# EventList.Item Property (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

Returns an item from a collection. The  **Item** property is the default property for all collections. Read-only.


## Syntax

 _expression_. **Item**( **_Index_**)

 _expression_A variable that represents an  **EventList** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Required| **Integer**|The index number of the object in its collection.|

### Return Value

Event


## Remarks

When retrieving objects from a collection, you can omit  **Item** from the expression because it is the default property for all collections. The following statement is equivalent to the syntax example given above:


```
objRet = object(index)
```

