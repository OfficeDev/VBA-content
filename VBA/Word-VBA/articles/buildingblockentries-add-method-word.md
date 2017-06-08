---
title: BuildingBlockEntries.Add Method (Word)
keywords: vbawd10.chm36241509
f1_keywords:
- vbawd10.chm36241509
ms.prod: word
api_name:
- Word.BuildingBlockEntries.Add
ms.assetid: 09578906-ea6d-9475-e026-b9dc437f451b
ms.date: 06/08/2017
---


# BuildingBlockEntries.Add Method (Word)

Creates a new building block entry in a template and returns a  **[BuildingBlock](buildingblock-object-word.md)** object that represents the new building block entry.


## Syntax

 _expression_ . **Add**( **_Name_** , **_Type_** , **_Category_** , **_Range_** , **_Description_** , **_InsertOptions_** )

 _expression_ An expression that returns a **[BuildingBlockEntries](buildingblockentries-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|Specifies the name of the building block entry. Corresponds to the  **[Name](buildingblock-name-property-word.md)** property of the **BuildingBlock** object.|
| _Type_|Required| **[WdBuildingBlockTypes](wdbuildingblocktypes-enumeration-word.md)**|Specifies the type of building block to create. Corresponds to the  **[Type](buildingblock-type-property-word.md)** property of the **BuildingBlock** object.|
| _Category_|Required| **String**|Specifies the category of the new building block entry. Corresponds to the  **[Category](buildingblock-category-property-word.md)** property of the **BuildingBlock** object.|
| _Range_|Required| **[Range](range-object-word.md)**|Specifies the value of the buildling block entry. Corresponds to the  **[Value](buildingblock-value-property-word.md)** property of the **BuildingBlock** object.|
| _Description_|Optional| **Variant**|Specifies the description of the buildling block entry. Corresponds to the  **[Description](buildingblock-description-property-word.md)** property of the **BuildingBlock** object.|
| _InsertOptions_|Optional| **[WdDocPartInsertOptions](wddocpartinsertoptions-enumeration-word.md)**|Specifies whether the building block entry is inserted as a page, a paragraph, or inline. If omitted, the default value is  **wdInsertContent** . Corresponds to the **[InsertOptions](buildingblock-insertoptions-property-word.md)** property for the **BuildingBlock** object.|

### Return Value

BuildingBlock


## Example

The following example creates a new building block entry and adds it to the template attached to the active document, and than sets the value of the building block to the selected text.


```vb
Dim objTemplate As Template 
Dim objBB As BuildingBlock 
 
Set objTemplate = ActiveDocument.AttachedTemplate 
Set objBB = objTemplate.BuildingBlockEntries.Add("Author Name", _ 
 wdTypeCustomTextBox, "Custom", Selection.Range)
```


## See also


#### Concepts


[BuildingBlockEntries Collection](buildingblockentries-object-word.md)

