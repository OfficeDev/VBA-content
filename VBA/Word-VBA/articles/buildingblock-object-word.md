---
title: BuildingBlock Object (Word)
keywords: vbawd10.chm3107
f1_keywords:
- vbawd10.chm3107
ms.prod: word
api_name:
- Word.BuildingBlock
ms.assetid: 2558b89f-8552-bb71-fa40-101cab2635ba
ms.date: 06/08/2017
---


# BuildingBlock Object (Word)

Represents a building block in a template. A building block is pre-built content, similar to autotext, that may contain text, images, and formatting.


## Remarks

Each  **BuildingBlock** object is a member of the **[BuildingBlocks](buildingblocks-object-word.md)** and **[BuildingBlockEntries](buildingblockentries-object-word.md)** collections. Building blocks are stored in Microsoft Word templates. Therefore, to access the building blocks available for a document, you need to access an attached template. Built-in building blocks are stored in the template named "Building Blocks.dotx".

 Use the **[Item](buildingblocks-item-method-word.md)** method of the collection or the **BuildingBlocks** collection to return an individual building block. The following example accesses the first building block in the first template in the **[Templates](templates-object-word.md)** collection.




```
Dim objTemplate As Template 
Dim objBB As BuildingBlock 
 
Set objTemplate = Templates(1) 
Set objBB = objTemplate.BuildingBlockEntries.Item(1)
```


 **Note**  Depending on how you access the collection, the collection returned may change. For example, if you access a collection of building blocks with a type of  **wdTypeAutoText** with a category of "General", the returned collection may be different from the collection returned if you access a collection of building blocks with a type of **wdTypeAutoText** with a category of "Custom". It is also different from the collection returned if you access the collection of building blocks with a type of **wdTypeCustomAutoText** with a category of "General". Therefore, the first item in a collection accessed from the **BuildingBlockEntries** collection may be different from the first item in the collection accessed from the **BuildingBlocks** collection.

To create a new building block, you can use the  **Add** method for either the **BuildingBlockEntries** collection or the **BuildingBlocks** collection. However, the recommended way to create a new building block is by using the **[Add](buildingblockentries-add-method-word.md)** method for the **BuildingBlockEntries** collection. The following example adds the selected text to the watermarks building block gallery of the first template in the **[Templates](templates-object-word.md)** collection.




```
Dim objTemplate As Template 
Dim objBB As BuildingBlock 
 
Set objTemplate = Templates(1) 
 
Set objBB = objTemplate.BuildingBlockEntries _ 
 .Add(Name:="New Building Block Entry", _ 
 Type:=wdTypeWatermarks, _ 
 Category:="General", _ 
 Range:=Selection.Range)
```

Use the  **[Insert](buildingblock-insert-method-word.md)** method to insert a new building block into a document. The following example inserts the first building block in the first template into the active document at the Insertion Point.




```
Dim objTemplate As Template 
Dim objBB As BuildingBlock 
 
Set objTemplate = Templates(1) 
Set objBB = objTemplate.BuildingBlockEntries.Item(1) 
 
objBB.Insert Selection.Range
```

Use the  **[Delete](buildingblock-delete-method-word.md)** method to remove a building block from a template. The following example deletes the first building block from the first template in the **Templates** collection.




```
Dim objTemplate As Template 
 
Set objTemplate = Templates(1) 
 
objTemplate.BuildingBlockEntries(1).Delete
```

 Building blocks are organized by category and type. Use the **[BuildingBlockTypes](buildingblocktypes-object-word.md)** collection to access individual **[BuildingBlockType](buildingblocktype-object-word.md)** objects. Use the **[Categories](categories-object-word.md)** collection to access individual **[Category](buildingblock-category-property-word.md)** objects. Then use the **BuildingBlocks** propery to access the **BuildingBlocks** collection for a **Category** object. The following example prints the type and category names of all the building blocks in the first template to the **Immediate Window**. (This example assumes that the **Immediate Window** is visible.)




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

Each building block has properties that contain information that applies uniquely to it, such as  **[Name](buildingblock-name-property-word.md)**, **[Description](buildingblock-description-property-word.md)**, **[Type](buildingblock-type-property-word.md)**, and **[Value](buildingblock-value-property-word.md)**.

For more information about building blocks, see [Working with Building Blocks](http://msdn.microsoft.com/library/c32a8972-a6fc-bb66-b62a-039b88580b37%28Office.15%29.aspx).


## Methods



|**Name**|
|:-----|
|[Delete](buildingblock-delete-method-word.md)|
|[Insert](buildingblock-insert-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Application](buildingblock-application-property-word.md)|
|[Category](buildingblock-category-property-word.md)|
|[Creator](buildingblock-creator-property-word.md)|
|[Description](buildingblock-description-property-word.md)|
|[ID](buildingblock-id-property-word.md)|
|[Index](buildingblock-index-property-word.md)|
|[InsertOptions](buildingblock-insertoptions-property-word.md)|
|[Name](buildingblock-name-property-word.md)|
|[Parent](buildingblock-parent-property-word.md)|
|[Type](buildingblock-type-property-word.md)|
|[Value](buildingblock-value-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
