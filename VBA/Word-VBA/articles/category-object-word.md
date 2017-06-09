---
title: Category Object (Word)
keywords: vbawd10.chm2910
f1_keywords:
- vbawd10.chm2910
ms.prod: word
api_name:
- Word.Category
ms.assetid: 5485ae39-fbcf-b18f-b1f9-945e220ecd2a
ms.date: 06/08/2017
---


# Category Object (Word)

Represents an individual category of a building block type.


## Remarks

Microsoft Word uses types and categories to organize building blocks. Each building block type is represented by a  **[WdBuildingBlockTypes](wdbuildingblocktypes-enumeration-word.md)** constant. Each category is a unique string that a user defines. Word comes with two categories already defined: "General" and "Custom"; you can create additional categories as you need.

Use the  **[Type](category-type-property-word.md)** property to access the building block type associated with a specific category. Use the **[BuildingBlocks](category-buildingblocks-property-word.md)** property to access the collection of building blocks for a category. The following example prints the type and category names of all the building blocks in the first template to the **Immediate Window** . (This example assumes that the **Immediate Window** is visible.)




```vb
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
 Debug.Print vbTab &; objCat.Name 
 Next 
 End If 
Next
```

Use the  **Item** method of the **Categories** collection to access an exising category; to create a new category, use the **Add** method of the **BuildingBlockEntries** collection. Set the value of the Category parameter.

For more information about building blocks, see [Working with Building Blocks](http://msdn.microsoft.com/library/c32a8972-a6fc-bb66-b62a-039b88580b37%28Office.15%29.aspx).


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


