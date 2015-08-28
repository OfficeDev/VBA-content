
# SmartTagProperties.Item Property (Access)

 **Last modified:** July 28, 2015

The  **Item** property returns a specific member of a collection either by position or by index. Read-only **SmartTagProperty**.

## Syntax

 _expression_. **Item**( **_Index_**)

 _expression_A variable that represents a  **SmartTagProperties** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Required| **Variant**|An expression that specifies the position of a member of the collection referred to by the expression argument. If a numeric expression, theindex argument must be a number from 0 to the value of the collection's **Count** property minus 1. If a string expression, theindex argument must be the name of a member of the collection|

## Remarks

If the value provided for the index argument doesn't match any existing member of the collection, an error occurs.

The  **Item** property is the default member of a collection, so you don't have to specify it explicitly. For example, the following two lines of code are equivalent:




```
Debug.Print Modules(0)
```




```
Debug.Print Modules.Item(0)
```


## See also


#### Concepts


 [SmartTagProperties Collection](1b2c7a0d-4dd3-e2b3-b44f-2a99994395f6.md)
#### Other resources


 [SmartTagProperties Object Members](9062e4fb-3df9-b474-3daa-776b220e2c34.md)
