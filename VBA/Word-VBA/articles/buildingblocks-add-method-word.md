---
title: BuildingBlocks.Add Method (Word)
ms.prod: word
api_name:
- Word.BuildingBlocks.Add
ms.assetid: 22725f33-4de0-95cd-d4a5-a2379b0130c4
ms.date: 06/08/2017
---


# BuildingBlocks.Add Method (Word)

Creates a new building block and returns a  **BuildingBlock** object.


## Syntax

 _expression_ . **Add**( **_Name_** , **_Range_** , **_Description_** , **_InsertOptions_** )

 _expression_ An expression that returns a **BuildingBlocks** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|Specifies the name of the building block entry. Corresponds to the  **[Name](buildingblock-name-property-word.md)** property of the **BuildingBlock** object.|
| _Range_|Required| **Range**|Specifies the value of the buildling block entry. Corresponds to the  **[Value](buildingblock-value-property-word.md)** property of the **BuildingBlock** object.|
| _Description_|Optional| **Variant**|Specifies the description of the buildling block entry. Corresponds to the  **[Description](buildingblock-description-property-word.md)** property of the **BuildingBlock** object.|
| _InsertOptions_|Optional| **[WdDocPartInsertOptions](wddocpartinsertoptions-enumeration-word.md)**|Specifies whether the building block entry is inserted as a page, a paragraph, or inline. If omitted, the default value is  **wdInsertContent** . Corresponds to the **[InsertOptions](buildingblock-insertoptions-property-word.md)** property for the **BuildingBlock** object.|

### Return Value

BuildingBlock


## Example

The following example adds a new building block auto text entry to the first template in the collection of templates.


```vb
Dim objTemplate As Template 
 
Set objTemplate = Templates(1) 
 
objTemplate.BuildingBlockTypes(wdTypeAutoText) _ 
 .Categories("General").BuildingBlocks _ 
 .Add Name:="New Building Block", _ 
 Range:=Selection.Range
```


## See also


#### Concepts


[BuildingBlocks Collection](buildingblocks-object-word.md)

