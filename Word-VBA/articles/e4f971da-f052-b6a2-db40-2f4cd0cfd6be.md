
# BuildingBlockType Object (Word)

 **Last modified:** July 28, 2015

Represents a type of building block. Each  **BuildingBlockType** object is a member of the ** [BuildingBlockTypes](fb179437-b736-dd99-3aea-125346aa7a3d.md)** collection.

## Remarks

Microsoft Word uses types and categories to organize building blocks. Each building block type is represented by a  ** [WdBuildingBlockTypes](be7fcedb-04fd-f27d-8f36-3120ca263f06.md)** constant. Use the ** [Categories](0daaeb0b-e6c8-488c-d965-bfdc4653d7e2.md)** property to access categories for a specific building block type. The following example prints the type and category names of all the building blocks in the first template to the **Immediate Window**. (This example assumes that the  **Immediate Window** is visible.)


```
Dim objTemplate As Template 
Dim objBBT As BuildingBlockType 
Dim objCat As Category 
Dim intCount As Integer 
Dim intCountCat As Integer 
 
Set objTemplate = Templates(1) 
 
For intCount = 1 To objTemplate.BuildingBlockTypes.Count 
 Set objBBT = objTemplate.BuildingBlockTypes(intCount) 
 If objBBT.Categories.Count > 0 Then 
 Debug.Print objBBT.Name 
 For intCountCat = 1 To objBBT.Categories.Count 
 Set objCat = objBBT.Categories(intCountCat) 
 Debug.Print vbTab &amp; objCat.Name 
 Next 
 End If 
Next
```

For more information about building blocks, see  [Working with Building Blocks](c32a8972-a6fc-bb66-b62a-039b88580b37.md).


## See also


#### Concepts


 [Word Object Model Reference](be452561-b436-bb9b-6f94-3faa9a74a6fd.md)
#### Other resources


 [BuildingBlockType Object Members](08b29414-6130-75b6-d3ed-77c2fd22b6b2.md)
