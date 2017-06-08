---
title: BuildingBlockTypes Object (Word)
keywords: vbawd10.chm2896
f1_keywords:
- vbawd10.chm2896
ms.prod: word
api_name:
- Word.BuildingBlockTypes
ms.assetid: fb179437-b736-dd99-3aea-125346aa7a3d
ms.date: 06/08/2017
---


# BuildingBlockTypes Object (Word)

Represents a collection of  **[BuildingBlockType](buildingblocktype-object-word.md)** objects.


## Remarks

Building block types are represented by  **[WdBuildingBlockTypes](wdbuildingblocktypes-enumeration-word.md)** constants. Use the **[Item](buildingblocktypes-item-method-word.md)** method to access a specific type in the **BuildingBlockTypes** collection.

To loop through the different building block types, use a  **For** loop with the **[Count](buildingblocktypes-count-property-word.md)** property. The following example loops through the building block types and prints the name in the **Immediate Window**. (This example assumes that the  **Immediate Window** is visible.)




```vb
Dim objTemplate As Template 
Dim intCount As Integer 
Dim objBBT As BuildingBlockType 
 
Set objTemplate = Templates(1) 
 
For intCount = 1 To objTemplate.BuildingBlockTypes.Count 
 Set objBBT = objTemplate.BuildingBlockTypes(intCount) 
 Debug.Print objBBT.Name 
Next
```

For more information about building blocks, see [Working with Building Blocks](http://msdn.microsoft.com/library/c32a8972-a6fc-bb66-b62a-039b88580b37%28Office.15%29.aspx).


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


