---
title: BuildingBlockType Object (Word)
keywords: vbawd10.chm2554
f1_keywords:
- vbawd10.chm2554
ms.prod: word
api_name:
- Word.BuildingBlockType
ms.assetid: e4f971da-f052-b6a2-db40-2f4cd0cfd6be
ms.date: 06/08/2017
---


# BuildingBlockType Object (Word)

Represents a type of building block. Each  **BuildingBlockType** object is a member of the **[BuildingBlockTypes](buildingblocktypes-object-word.md)** collection.


## Remarks

Microsoft Word uses types and categories to organize building blocks. Each building block type is represented by a  **[WdBuildingBlockTypes](wdbuildingblocktypes-enumeration-word.md)** constant. Use the **[Categories](buildingblocktype-categories-property-word.md)** property to access categories for a specific building block type. The following example prints the type and category names of all the building blocks in the first template to the **Immediate Window** . (This example assumes that the **Immediate Window** is visible.)


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

For more information about building blocks, see [Working with Building Blocks](http://msdn.microsoft.com/library/c32a8972-a6fc-bb66-b62a-039b88580b37%28Office.15%29.aspx).


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


