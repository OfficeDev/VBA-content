---
title: Create a Custom Building Block Gallery
ms.prod: word
ms.assetid: 472688b6-205c-c88d-5a7e-26334ec5eeeb
ms.date: 06/08/2017
---


# Create a Custom Building Block Gallery

A building block gallery is a collection of building blocks that are of the same type. There are 32 different types of building block collections that you can create (by using the  **[WdBuildingBlockTypes](wdbuildingblocktypes-enumeration-word.md)** enumeration). Each of these types is a gallery. Word provides some of these galleries as built-in building block collections, but most of them are provided so that you can organize your own building blocks. 

To provide added flexibility with your custom galleries, you can create categories within each gallery to further organize your custom building blocks.

The objects used in this sample are:

-  **[Template](template-object-word.md)**
    
-  **[Range](range-object-word.md)**
    
-  **[BuildingBlockEntries](buildingblockentries-object-word.md)**
    
-  **[BuildingBlock](buildingblock-object-word.md)**
    

## Sample

The following example creates a building block gallery that contains three building blocks.


 **Note**  This example assumes that there are at least three paragraphs in the document.


```vb
Sub CreateBuildingBlockGallery() 
 Dim objTemplate As Template 
 Dim conType As WdBuildingBlockTypes 
 Dim objRange As Range 
 
 Set objTemplate = ActiveDocument.AttachedTemplate 
 conType = wdTypeCustom1 
 
 Set objRange = ActiveDocument.Paragraphs(1).Range 
 objTemplate.BuildingBlockEntries.Add _ 
 "Heading 1", conType, "Main Headings", objRange 
 
 Set objRange = ActiveDocument.Paragraphs(2).Range 
 objTemplate.BuildingBlockEntries.Add _ 
 "Heading 2", conType, "Secondary Headings", objRange 
 
 Set objRange = ActiveDocument.Paragraphs(3).Range 
 objTemplate.BuildingBlockEntries.Add _ 
 "Heading 3", conType, "Tertiary Headings", objRange 
End Sub
```


## Additional Resources


-  [Working with Building Blocks](working-with-building-blocks.md)
    

