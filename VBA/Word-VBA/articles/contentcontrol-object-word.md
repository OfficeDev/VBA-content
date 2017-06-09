---
title: ContentControl Object (Word)
keywords: vbawd10.chm4067
f1_keywords:
- vbawd10.chm4067
ms.prod: word
api_name:
- Word.ContentControl
ms.assetid: 783dec26-9b63-11f8-6187-985f9c815f27
ms.date: 06/08/2017
---


# ContentControl Object (Word)

An individual content control. Content controls are bounded and potentially labeled regions in a document that serve as containers for specific types of content. Individual content controls may contain contents such as dates, lists, or paragraphs of formatted text. The  **ContentControl** object is a member of the **[ContentControls](contentcontrols-object-word.md)** collection.


## Remarks

Use the  **[Add](contentcontrols-add-method-word.md)** method of the **ContentControls** collection to create a content control. Use the Type parameter of the **Add** method to specify the type of content control to create. The following example create a new drop-down list content control and adds several items to the list.


```vb
Dim objCC As ContentControl 
 
Set objCC = ActiveDocument.ContentControls.Add(Type:=wdContentControlDropdownList) 
 
'List entries 
objCC.DropdownListEntries.Add "Cat" 
objCC.DropdownListEntries.Add "Dog" 
objCC.DropdownListEntries.Add "Horse" 
objCC.DropdownListEntries.Add "Monkey" 
objCC.DropdownListEntries.Add "Snake" 
objCC.DropdownListEntries.Add "Other"
```

Use the  **Type** property to change the content control to a different type of content control. For example, perhaps you want to change from a date control to a text control. However, you may not be able to change all content controls to another type; some may not allow changing their type. In addition, depending on the contents of a content control, you may not be able to change the type. For example, if the content control that you want to change to does not allow the type of content that is in the existing content control, attempting to change the type is not allowed and generates a run-time error.

The following example inserts a date content control and sets the value of the control, and then changes the control to a text content control.




```vb
Dim objCC As ContentControl 
 
Set objCC = ActiveDocument.ContentControls.Add(wdContentControlDate) 
 
objCC.Range.Text = "January 1, 2007" 
objCC.Type = wdContentControlText
```

Use the  **SetPlaceholderText** method to change the placeholder text from the default string to something more appropriate for the control. Use the **Title** property to specify the title text for the control. This displays above the control when the cursor is positioned inside the control or the mouse pointer is positioned over the control.

Depending on the type of content control that you have, you may not be able to use all the properties and methods of the  **ContentControl** object.

Not all content control properties apply to all the different types of content controls. The following table lists which properties apply to which types of content controls.



|**Property/Method**|**Applies To**|
|:-----|:-----|
| **[BuildingBlockCategory](contentcontrol-buildingblockcategory-property-word.md)** property|BuildingBlock Gallery content controls (wdContentControlBuildingBlockGallery)|
| **[BuildingBlockType](contentcontrol-buildingblocktype-property-word.md)** property|BuildingBlock Gallery content controls (wdContentControlBuildingBlockGallery)|
| **[DateDisplayFormat](contentcontrol-datedisplayformat-property-word.md)** property|Date content controls (wdContentControlDate)|
| **[DateDisplayLocale](contentcontrol-datedisplaylocale-property-word.md)** property|Date content controls (wdContentControlDate)|
| **[DateStorageFormat](contentcontrol-datestorageformat-property-word.md)** property|Date content controls (wdContentControlDate)|
| **[DropdownListEntries](contentcontrol-dropdownlistentries-property-word.md)** property|Combo box and drop-down list content controls (wdContentControlComboBox and wdContentControlDropdownList)|
| **[MultiLine](contentcontrol-multiline-property-word.md)** property|Plain text content controls (wdContentControlText)|
| **[Ungroup](contentcontrol-ungroup-method-word.md)** method|Group content controls (wdContentControlGroup)|

## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

