
# CommandBars.Item Property (Office)

Gets a  **CommandBar** object from the **CommandBars** collection. Read-only.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **Item**( ** _Index_** )

 _expression_ Required. A variable that represents a **[CommandBars](0e312e21-14ee-5055-d604-b66e61c53b47.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Variant**|The name or index number of the object to be returned.|

## Example

Item is the default member of the object or collection. The following two statements both assign a CommandBar object to cmdBar.


```vb
Set cmdBar = CommandBars.Item("Standard") 
Set cmdBar = CommandBars("Standard")
```


## See also


#### Concepts


[CommandBars Object](0e312e21-14ee-5055-d604-b66e61c53b47.md)
#### Other resources


[CommandBars Object Members](c11db22d-b7bb-20a2-a455-e441cb8d5bc0.md)