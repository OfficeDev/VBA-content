
# SharedResources.Item Property (Access)

 **Last modified:** July 28, 2015

The  **Item** property returns a specific member of a collection either by position or by index. Read-only **Object**.

## Syntax

 _expression_. **Item**( **_Index_**)

 _expression_A variable that represents a  **SharedResources** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Required| **Long**||

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


 [SharedResources Collection](45323141-e7df-1c70-efe2-926c1990d5e0.md)
#### Other resources


 [SharedResources Object Members](3dfef725-97ed-5a11-3b28-3458f2772f32.md)
