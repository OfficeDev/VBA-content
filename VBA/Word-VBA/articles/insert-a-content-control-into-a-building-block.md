---
title: Insert a Content Control into a Building Block
ms.prod: word
ms.assetid: f6e917d9-f756-e36e-696f-bc7cf84b92e3
ms.date: 06/08/2017
---


# Insert a Content Control into a Building Block

Building blocks and content controls are extremely flexible. You can create building blocks that contain content controls, or you can create content controls that use building blocks to present rich content selections to the user. This topic shows how to insert a content control into a building block, which users can then insert into their documents.

The objects used in this sample are:

-  **[Template](template-object-word.md)**
    
-  **[BuildingBlock](buildingblock-object-word.md)**
    
-  **[Range](range-object-word.md)**
    
-  **[ContentControl](contentcontrol-object-word.md)**
    
The following code inserts a content control into the active document, and then adds the content control to the collection of building blocks in the template attached to the active document.



```vb
Sub InsertContentControlIntoBuildingBlock() 
 Dim objCC As ContentControl 
 Dim objBB As BuildingBlock 
 Dim objTemplate As Template 
 Dim objRange As Range 
 
 Set objTemplate = ActiveDocument.AttachedTemplate 
 Set objCC = ActiveDocument.Range.ContentControls _ 
 .Add(wdContentControlComboBox) 
 
 objCC.DropdownListEntries.Add "Outstanding" 
 objCC.DropdownListEntries.Add "Good" 
 objCC.DropdownListEntries.Add "Fair" 
 
 Set objRange = ActiveDocument.Range 
 Set objBB = objTemplate.BuildingBlockEntries.Add("OGF Rating Scale", _ 
 wdTypeCustom1, "Ratings", objRange) 
End Sub
```


## Additional Resources


-  [Working with Building Blocks](working-with-building-blocks.md)
    
-  [Working with Content Controls](working-with-content-controls.md)
    

