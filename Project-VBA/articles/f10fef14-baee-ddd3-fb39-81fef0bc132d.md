
# Shapes.Value Property (Project)
Gets an individual  **Shape** object in the **Shapes** collection. Read-only **Shape**.

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Property value](#sectionSection2)


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Value**

 _expression_A variable that represents a  **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Required| **Variant**|Can be a  **String** value for the name of the shape or a **Long** value for the ordinal index number of the shape.|

## Remarks
<a name="sectionSection1"> </a>

 **Value** is the default property for the **Shapes** object. For example, create a report namedTable Tests that contains a table. The following statement in the **Immediate** window of the VBE prints the name of the table.


```
? ActiveProject.Reports("Table Tests").Shapes.Value(1).Name
```

If you leave off the  **Shapes** property, the following statement is effectively the same as the previous statement.




```
? ActiveProject.Reports("Table Tests").Shapes(1).Name
```

 **Shapes.Item** acts like **Shapes.Value**, except  **Item** is a method:




```
? ActiveProject.Reports("Table Tests").Shapes.Item(1).Name
```


## Property value
<a name="sectionSection2"> </a>

 **SHAPE**


## See also
<a name="sectionSection2"> </a>


#### Other resources


 [Shapes Object](6e42040c-dd5a-de4c-afa8-f9e33d1e5054.md)
 [Item Method](43fba4f4-f3d3-20a0-2c77-15e31dcdcbf5.md)
