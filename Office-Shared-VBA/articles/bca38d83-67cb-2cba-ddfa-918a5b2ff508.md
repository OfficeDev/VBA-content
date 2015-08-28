
# CommandBars.Item Property (Office)

 **Last modified:** July 28, 2015

 **In this article**
 [](#sectionSection0)
 [Syntax](#sectionSection1)
 [Example](#sectionSection2)


Gets a  **CommandBar** object from the **CommandBars** collection. Read-only.


## 
<a name="sectionSection0"> </a>


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax
<a name="sectionSection1"> </a>

 _expression_. **Item**( **_Index_**)

 _expression_Required. A variable that represents a  ** [CommandBars](0e312e21-14ee-5055-d604-b66e61c53b47.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Required| **Variant**|The name or index number of the object to be returned.|

## Example
<a name="sectionSection2"> </a>

Item is the default member of the object or collection. The following two statements both assign a CommandBar object to cmdBar.


```
Set cmdBar = CommandBars.Item("Standard") 
Set cmdBar = CommandBars("Standard")
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [CommandBars Object](0e312e21-14ee-5055-d604-b66e61c53b47.md)
#### Other resources


 [CommandBars Object Members](c11db22d-b7bb-20a2-a455-e441cb8d5bc0.md)
