---
title: Filter the Building Block List in a Content Control
ms.prod: word
ms.assetid: 0eb233f4-f024-27dd-05d0-4f49c26d1bbc
ms.date: 06/08/2017
---


# Filter the Building Block List in a Content Control

When you use content controls together with building blocks, you can help users by indicating what content they can insert and by limiting their choices. You can do this is by filtering the list of building blocks that are shown for a content control. To filter a building block list for a content control, you use the  **[BuildingBlockType](contentcontrol-buildingblocktype-property-word.md)** property for the content control. You can further filter the list of building blocks by setting the **[BuildingBlockCategory](contentcontrol-buildingblockcategory-property-word.md)** property for the content control.

You can filter the list of building blocks for a content control when you create the content control. However, you can also programmatically filter the list of building blocks based on the position of the cursor or on the value of another content control. To start, you need a custom building block gallery. To learn how to create a custom building block gallery, see  [Create a Custom Building Block Gallery](create-a-custom-building-block-gallery.md).




-  **[BuildingBlock](buildingblock-object-word.md)**
    
-  **[ContentControl](contentcontrol-object-word.md)**
    

## Sample 1

The following example shows how to filter a building block list to a specific gallery.


 **Note**  Run the code in the "Create a Custom Building Block Gallery" topic before running the code in this topic.


```vb
Sub CreateFilteredContentControl_SampleOneA() 
 Dim objCC As ContentControl 
 
 Set objCC = ActiveDocument.ContentControls.Add( _ 
 wdContentControlBuildingBlockGallery, Selection.Range) 
 
 objCC.BuildingBlockType = wdTypeCustom1 
End Sub
```

You can filter a building block list even further by specifying a specific category within the gallery. The following example shows how to filter a building block list to include only the building blocks within a category of a specified gallery.




```vb
Sub CreateFilteredContentControl_SampleOneB() 
 Dim objCC As ContentControl 
 
 Set objCC = ActiveDocument.ContentControls.Add( _ 
 wdContentControlBuildingBlockGallery, Selection.Range) 
 
 objCC.BuildingBlockType = wdTypeCustom1 
 objCC.BuildingBlockCategory = "Tertiary Headings" 
End Sub
```


## Sample 2

To filter a building block list based on the position of the cursor, you need to use the  **ContentControlOnEnter** event. For example, if you have a content control named Report Type that can be set to "financial" or "marketing", you can have a building block content control that shows a list of possible disclaimers. The content control for the disclaimers would show all disclaimers if the Report Type is not set, and only the appropriate subset if the property is set. The following example filters the list of building blocks for a content control based on the value of another content control in the document.


```vb
Private Sub Document_ContentControlOnEnter(ByVal ContentControl As ContentControl) 
 Dim objCC As ContentControl 
 Dim objType As ContentControl 
 
 Set objCC = ContentControl 
 Set objType = ActiveDocument.ContentControls.Item("Report Type") 
 
 If objCC.Title = "Disclaimer" Then 
 Select Case objType.Range.Text 
 Case "Financial" 
 objCC.BuildingBlockType = wdTypeCustom1 
 objCC.BuildingBlockCategory = "Financial Disclaimers" 
 
 Case "Marketing" 
 objCC.BuildingBlockType = wdTypeCustom1 
 objCC.BuildingBlockCategory = "Marketing Disclaimers" 
 
 End Select 
 End If 
End Sub
```


## Additional Resources


-  [Working with Building Blocks](working-with-building-blocks.md)
    
-  [Working with Content Controls](working-with-content-controls.md)
    

