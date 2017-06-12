---
title: Working with Building Blocks
ms.prod: word
ms.assetid: c32a8972-a6fc-bb66-b62a-039b88580b37
ms.date: 06/08/2017
---


# Working with Building Blocks

Introduced in Word 2007, building blocks are similar to autotext entries in previous versions. However, with building blocks, you can organize entries in a way that enables you to determine how a user uses them. A custom building block helps users insert rich content anywhere in a document by using a formatted drop-down list. When used together with content controls, building blocks can enable you to develop powerful solutions quickly and easily.

The building blocks object model includes three new objects and four new collections. These enable you to create an organizational structure that works for your specific needs and to modify the structure for a specific solution. The new objects and collections are listed in the following table.


|**Name**|**Description**|
|:-----|:-----|
| **[BuildingBlock](buildingblock-object-word.md)**|A specific building block entry.|
| **[BuildingBlocks](buildingblocks-object-word.md)**|A collection of building block entries in a template that are of the same type and category.|
| **[BuildingBlockEntries](buildingblockentries-object-word.md)**|A collection of all the building blocks in a template.|
| **[BuildingBlockType](buildingblocktype-object-word.md)**|A building block type. |
| **[BuildingBlockTypes](buildingblocktypes-object-word.md)**|A collection of building block types.|
| **[Category](category-object-word.md)**|A building block category.|
| **[Categories](categories-object-word.md)**|A collection of building block categories.|

## Understanding Building Blocks

Building blocks are organized by type and category. Building block types are composed of a limited number of  **[WdBuildingBlockTypes](wdbuildingblocktypes-enumeration-word.md)** constants. Although there are a limited number of these constants, that number is not small. There are 35 different **WdBuildingBlockTypes** constants. These types help you to define and organize your building blocks and, although you cannot create additional building block types, you can create an unlimited number of categories for each type.

Categories are composed of an unlimited number of strings that you can define to organize your custom building blocks. Building blocks are stored in templates. By default, the templates that are included with Word have building block categories like "General" and "Built-In". However, you are not limited to just the categories that are included in these templates. A category can be any string that you define. Types and categories are explained later in this topic.

Because you can organize building blocks into types and categories, building blocks can be incredibly flexible. For example, you can have a building block named "Title" that has a type of  **wdTypeBibliography** and a category of "Book Titles" and another building block named "Title" that has a type of **wdTypeBibliography** and a category of "Movie Titles" and then you can have yet another building block named "Title" that has a type of **wdTypeCustomHeaders** and a category of "Book Titles" and so on. The incredible flexibility that building blocks provide help you to create custom solutions without writing code.

However, building block are also programmable. You can create new building blocks, delete building blocks, and manage building blocks programmatically. You can also watch for when users insert new building blocks into a document by using the  **[BuildingBlockInsert](document-buildingblockinsert-event-word.md)** event. Plus, you can use building blocks with content controls to give you even greater control over which building blocks users can insert into their documents. For example, you can use a building block content control to filter the types of building blocks that a user sees, which means that the user cannot insert a building block into a document that is not allowed at a specific place in the document. There are several examples in the following sections that show you how to use the building block objects to work with building blocks programmatically.


## Simple Tasks

The following sections provide simple examples of how to do specific tasks using the building block objects. You can find additional code examples in the object topics and in many of the member topics.


## Creating a Custom Building Block

Creating a custom building block is as simple as using the  **[Add](buildingblocks-add-method-word.md)** method for the **BuildingBlockEntries** collection. You can also use the **[Add](buildingblockentries-add-method-word.md)** method for the **BuildingBlocks** collection; however, this method may raise a run-time error if there are currently no building blocks for the specified type or category. As explained in the table of objects, the **BuildingBlocks** collection is a collection of building blocks for a specific type and category. The **BuildingBlocksEntries** collection contains all the building blocks for a template. Therefore, the preferred way to add new building blocks programmatically is to use the **Add** method for the **BuildingBlockEntries** collection.

The following code example collapses the current selection, creates a range and specifies the text for the range, and then adds the selection as a custom building block to the collection of building block entries in the template attached to the current document.




```vb
Sub AddCustomBuildingBlock() 
 
 Dim objTemplate As Template 
 Dim objBB As BuildingBlock 
 Dim objRange As Range 
 
 ' Set the template to store the building block 
 Set objTemplate = ActiveDocument.AttachedTemplate 
 
 ' Collapse the range, set the range, and add the text 
 Selection.Collapse 
 Set objRange = Selection.Range 
 objRange.Text = "Building blocks for the technically challenged" 
 
 ' Add the building block to the template 
 Set objBB = objTemplate.BuildingBlockEntries.Add( _ 
 Name:="Title", _ 
 Type:=wdTypeCustomHeaders, _ 
 Category:="Book Titles", _ 
 Range:=objRange) 
 
End Sub
```


## Adding a New Category

As mentioned previously, you can add an unlimited number of categories. However, there is no  **Add** method for the **Categories** collection. Therefore, to add a new category to the collection, you need to add a new building block. For example, in the previous code sample, if the "Book Titles" category does not exist when you run the code, Word adds it to the **Categories** collection.


## Accessing an Existing Building Block

At some point you will want to access one of the building blocks that you have, whether that is a custom building block or one of the built-in building blocks. You could use the  **BuildingBlockEntries** collection; however, because building blocks can share the same name, you would need to identify the type and category for the building block before knowing which one you want returned. Therefore, the best way to access existing building blocks is through the **BuildingBlocks** collection.

The following code example accesses the building block that you added in the previous code example.




```vb
Sub GetExistingBuildingBlock() 
 
 Dim objTemplate As Template 
 Dim objBB As BuildingBlock 
 
 ' Set the template to store the building block 
 Set objTemplate = ActiveDocument.AttachedTemplate 
 
 ' Access the building block through the type and category 
 Set objBB = objTemplate.BuildingBlockTypes(wdTypeCustomHeaders) _ 
 .Categories("Book Titles").BuildingBlocks("Title") 
 
End Sub
```


## Inserting a Building Block into a Document

After you have access to a building block, use the  **Insert** method of the **BuildingBlock** object to insert it into a document. The following code example expands the previous code sample by adding a line for inserting the building into the active document at the Insertion Point (or for replacing the selected text, if text is selected).


 **Note**  When you insert a building block by using the ribbon, Word automatically determines certain things about the building block, such as where to insert it; however, when you insert a building block through the object model, none of this built-in intelligence automatically happens. For example, when you insert a header building block by using the ribbon, Word automatically determines to replace the existing header. When inserting the same header building block by using the object model, you need to explicitly specify where to place the building block text.


```vb
Sub InsertExistingBuildingBlock() 
 
 Dim objTemplate As Template 
 Dim objBB As BuildingBlock 
 
 ' Set the template to store the building block 
 Set objTemplate = ActiveDocument.AttachedTemplate 
 
 ' Access the building block through the type and category 
 Set objBB = objTemplate.BuildingBlockTypes(wdTypeCustomHeaders) _ 
 .Categories("Book Titles").BuildingBlocks("Title") 
 
 ' Insert the building block into the document replacing any selected text. 
 objBB.Insert Selection.Range 
 
End Sub
```


## Filtering a List of Building Blocks in a Content Control

If you combine building blocks with content controls, you can filter which building blocks a user can access. You do this using a content control and an event. When a user enters a content control, the  **ContentControlOnEnter** event for the **Document** object fires. This event has a parameter for the active content control. You can determine whether the content control is a building block content control. If it is, you use the **BuildingBlockType** property and the **BuildingBlockCategory** property to identify which type and category to use to filter the list of building blocks that are available for the content control. This specifies which building blocks show up in the drop-down list in the content control header.

The following code example assumes that there is at least one content control in the document. If the content control is a building block content control, the list of building blocks displayed in the building block list in the content control header includes only those added by using the AddCustomBuildingBlock subroutine shown earlier in this topic. For more information about content controls, see  [Working with Content Controls](working-with-content-controls.md).




```vb
Private Sub Document_ContentControlOnEnter(ByVal ContentControl As ContentControl) 
 Dim objCC As ContentControl 
 
 Set objCC = ContentControl 
 
 If objCC.Type = wdContentControlBuildingBlockGallery Then 
 objCC.BuildingBlockType = wdTypeCustomHeaders 
 objCC.BuildingBlockCategory = "Book Titles" 
 End If 
End Sub
```


