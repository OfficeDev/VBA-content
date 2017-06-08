---
title: Set the Building Blocks That You Can Use for a Content Control
ms.prod: word
ms.assetid: 6723a4c4-f96c-7bbd-a978-66602ab693c7
ms.date: 06/08/2017
---


# Set the Building Blocks That You Can Use for a Content Control

A document building block is a predesigned piece of content, such as a cover page or a header or footer. Word includes a library of document building blocks that users can choose from to insert into a document. 

A  [ContentControl Object (Word)](contentcontrol-object-word.md) object with a [ContentControl.Type Property (Word)](contentcontrol-type-property-word.md) property value of **wdContentControlBuildingBlockGallery** specifies a content control that can contain document building blocks.

The  **[WdBuildingBlockTypes](wdbuildingblocktypes-enumeration-word.md)** enumeration contains each building block type. You can only use the following building block types within a building block gallery content control:


- AutoText
    
- Tables
    
- Equations
    
- Quick Parts
    
- Custom 1 though Custom 5
    
- Custom AutoText
    
- Custom Tables
    
- Custom Equations
    
- Custom Quick Parts
    
For more information about content controls, see  [Working with Content Controls](working-with-content-controls.md).
The objects used in this sample are:

-  **[ContentControl](contentcontrol-object-word.md)**
    
-  **[ContentControls](contentcontrols-object-word.md)**
    

## Sample

The following code sample instantiates a building block gallery content control and then adds a building block to the content control.


```vb
Sub SetBuildingBlock()
 
    Dim strTitle As String
    strTitle = "My Equation"
    Dim objContentControl As ContentControl
     
    Set objContentControl = ActiveDocument.ContentControls _
        .Add(wdContentControlBuildingBlockGallery)
    objContentControl.Title = strTitle
    objContentControl.BuildingBlockType = wdTypeEquations
   
End Sub
```


