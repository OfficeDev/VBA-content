---
title: Working with Content Controls
ms.prod: word
ms.assetid: b4092c71-a383-f1db-8d68-de69e8d8a86b
ms.date: 06/08/2017
---


# Working with Content Controls

## What Are Content Controls?

Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls can contain content such as dates, lists, or paragraphs of formatted text. In some cases, content controls might remind you of forms. However, they are much more powerful, flexible, and useful because they enable you to create rich, structured blocks of content. Content controls enable you to author templates that insert well-defined blocks into your documents. Content controls enable you to:


- Specify structured regions in a template. Each structured region has its own unique ID so that you can read from and write to it. Examples of types of structured regions (or content controls) are combo boxes, pictures, text blocks, and calendars.
    
- Determine the behavior of content controls. Each content control takes up a portion of a document and, as the template author, you can specify what each region does. For example, if you want a region of your template to be a calendar, you insert a calendar content control in that area of the document, which automatically determines what that block of content does. Similarly, if you want a section of a template to display an image, create a picture content control in that area. In this way, you can build a template with predefined block types.
    
- Restrict the ability to modify content controls within a document. Each content control can be restricted, so that it cannot be deleted or edited. This is useful if, for example, you have copyright information in a template that the user should be able to read but not edit. Or, as another example, you can also lock a content control that you have placed within a template document so that a user does not accidentally delete the content contained in the content control. This makes templates more robust than in previous versions.
    
- Map the contents of a content control to data in a custom XML part. For example, if you insert plain text content controls into cells of a table of stock prices, you can map the content controls in the table cells to nodes in an XML file that contain the current stock prices. When the prices change, an add-in can programmatically update the attached XML file, which is bound to each plain text content control, and the new, updated prices automatically appear in the table.
    
The easiest way to create a content control is through the user interface (although you can also create them programmatically). To create a content control through the user interface (UI), select the content that you want to turn into a content control (for example, some text or a picture) and then choose the content control type you want from the content controls section of the Developer ribbon. This creates a content control around the selected content.


## Content Controls in the Word Object Model

The following table shows the objects in the Word object model that relate to content controls.



|**Name**|**Description**|
|:-----|:-----|
| **[ContentControl](contentcontrol-object-word.md)**|Each  **ContentControl** object represents an individual content control within a document. Use the **ContentControls** collection to access individual **ContentControl** objects.|
| **[ContentControls](contentcontrols-object-word.md)**|You can use the  **ContentControls** properties of the **[Document](document-object-word.md)**,  **[Range](range-object-word.md)**, and  **[Selection](selection-object-word.md)** objects to access the collection of content controls. You can also use the **[SelectContentControlsByTitle](document-selectcontentcontrolsbytitle-method-word.md)** method and the **[SelectContentControlsByTag](document-selectcontentcontrolsbytag-method-word.md)** method of the **Document** object to access a **ContentControls** collection that includes specific content controls that all have the same title or tag value.|
| **[ContentControlListEntry](contentcontrollistentry-object-word.md)**|When a content control is a drop-down list or combo box, the  **ContentControlListEntry** object represents individual items within the list.|
| **[ContentControlListEntries](contentcontrollistentries-object-word.md)**|Use the  **[DropdownListEntries](contentcontrol-dropdownlistentries-property-word.md)** property of the **ContentControl** object to access all the items in an individual drop-down list or combo box.|
Each of these objects or collections has methods and properties that allow you to work with the content controls both individually and as a collection. Because there are various types of content controls (see the following section "Types of Content Controls"), the  **ContentControl** object has members that might not apply to all the different types of content controls. The following table shows those properties and methods of the **ContentControl** object that only apply to certain types of content controls.


**Note**: For a complete list of all properties and methods of the **ContentControl** object, see [Content Controls](https://msdn.microsoft.com/en-us/library/bb157891.aspx).

|**Property/Method**|**Applies To**|
|:-----|:-----|
| **[BuildingBlockCategory](contentcontrol-buildingblockcategory-property-word.md)** property|BuildingBlock Gallery content controls (**wdContentControlBuildingBlockGallery**)|
| **[BuildingBlockType](contentcontrol-buildingblocktype-property-word.md)** property|BuildingBlock Gallery content controls (**wdContentControlBuildingBlockGallery**)|
| **[DateDisplayFormat](contentcontrol-datedisplayformat-property-word.md)** property|Date content controls (**wdContentControlDate**)|
| **[DateDisplayLocale](contentcontrol-datedisplaylocale-property-word.md)** property|Date content controls (**wdContentControlDate**)|
| **[DateStorageFormat](contentcontrol-datestorageformat-property-word.md)** property|Date content controls (**wdContentControlDate**)|
| **[DropdownListEntries](contentcontrol-dropdownlistentries-property-word.md)** property|Combo box and drop-down list content controls (**wdContentControlComboBox** and **wdContentControlDropdownList**)|
| **[MultiLine](contentcontrol-multiline-property-word.md)** property|Plain-text content controls (**wdContentControlText**)|
| **[Ungroup](contentcontrol-ungroup-method-word.md)** method|Group content controls (**wdContentControlGroup**)|
| **[SetCheckedSymbol](contentcontrol-setcheckedsymbol-method-word.md)** method|Check Box content control (**wdContentControlCheckBox**)|
| **[SetUncheckedSymbol](contentcontrol-setuncheckedsymbol-method-word.md)** method|Check Box content control (**wdContentControlCheckBox**)|

## Types of Content Controls

There are eight different types of content controls that you can add to a document, each of which is represented in a new enumeration called **[WdContentControlType](wdcontentcontroltype-enumeration-word.md)**.



|**Content Control Type**|**Description**|**WdContentControlType Constant**|
|:-----|:-----|:-----|
||A checkbox.| **wdContentControlCheckBox**|
|Calendar|A date-time picker.| **wdContentControlDate**|
|Building Block|Enables the user to choose from specified building blocks.| **wdContentControlBuildingBlockGallery**|
|Drop-Down List|A drop-down list.| **wdContentControlDropDownList**|
|Group|Defines a protected region of a document that users cannot edit or delete. A group control can contain any document items, such as text, tables, graphics, and other content controls.| **wdContentControlGroup**|
|Combo Box|A combo box.| **wdContentControlComboBox**|
|Picture|A picture.| **wdContentControlBlockPicture**|
|Rich Text|A block of rich text.| **wdContentControlRichText**|
|Plain Text|A block of plain text.| **wdContentControlText**|

## Content Control Events

In addition to the properties and methods available with the content control object model in Word, you can also use several events that allow you to run code when adding or removing a content control or when a user edits a content control. The following list describes each of the events and when the event code runs. All of these events are members of the **Document** object.



|**Event Name**|**Description**|
|:-----|:-----|
| **[ContentControlAfterAdd](document-contentcontrolafteradd-event-word.md)**|Occurs after adding a new content control to a document. This event runs whether the user adds the content control by using the tools in the UI or adds them by using code.|
| **[ContentControlBeforeContentUpdate](document-contentcontrolbeforecontentupdate-event-word.md)**|Occurs before Word updates the content in a content control.|
| **[ContentControlBeforeDelete](document-contentcontrolbeforedelete-event-word.md)**|Occurs before a user deletes a content control. This event runs whether the user deletes the content control by using the tools in the UI or deletes them by using code.|
| **[ContentControlBeforeStoreUpdate](document-contentcontrolbeforestoreupdate-event-word.md)**|Occurs before Word updates the contents of a content control from data in the document's data store.|
| **[ContentControlOnEnter](document-contentcontrolonenter-event-word.md)**|Occurs when a user enters a content control.|
| **[ContentControlOnExit](document-contentcontrolonexit-event-word.md)**|Occurs when a user exits a content control.|

## Working with the Code

Whether you want to add a content control, delete a content control, or access and manipulate existing content controls, you can do it with code. The following sections are just a few samples of what you can do.


## Adding a Content Control

As mentioned previously, there are eight different types of content controls that you can add to your documents. Use the  **Add** method of the **ContentControls** collection to add a content control to a document. The following example adds a date picker to the active document and sets the date value to the current date.


```vb
Sub AddDatePicker() 
 
    Dim objCC As ContentControl 
    Dim objDate As Date 
 
    Set objCC = ActiveDocument.ContentControls _ 
        .Add(wdContentControlDate) 
    objDate = Date 
    objCC.Range.Text = objDate 
     
End Sub
```

You can use the same basic construction to add any of the different types of content controls to a document.


## Adding a Title to a Content Control

Use the  **Title** property to add a title to a content control. This is text that users see, and it can help them to know what type of data to enter into the content control. The following example adds a new plain-text content control to the active document and sets the title, or display text, for the control.


```vb
Sub SetTitleForContentControl() 
 
    Dim objCC As ContentControl 
     
    Set objCC = ActiveDocument.ContentControls _ 
        .Add(wdContentControlText) 
         
    objCC.Title = "Please enter your name" 
     
End Sub
```


## Modifying Placeholder Text to a Content Control

Placeholder text is temporary text. It can be a simple one-word or two-word description (similar to the title) or it can be a more thorough description (such as numbered steps). Modifying the placeholder text is the same regardless of the type of content control or the expected contents of the content control. The following example adds a drop-down list to the active document, sets the placeholder text for the control, and then fills the list with the names of several animals.


```vb
Sub SetPlaceholderText() 
 
    Dim objCC As ContentControl 
     
    Set objCC = Selection.ContentControls _ 
        .Add(wdContentControlComboBox) 
    objCC.Title = "Favorite Animal" 
    objCC.SetPlaceholderText _ 
        Text:="Please select your favorite animal " 
     
    'List entries 
    objCC.DropdownListEntries.Add "Cat" 
    objCC.DropdownListEntries.Add "Dog" 
    objCC.DropdownListEntries.Add "Horse" 
    objCC.DropdownListEntries.Add "Monkey" 
    objCC.DropdownListEntries.Add "Snake" 
    objCC.DropdownListEntries.Add "Other" 
 
End Sub
```

These are just a few of the ways that you can use the object model to manipulate content controls in your documents. For more examples, see the  [How To](how-do-i-word-vba-reference.md) section.


