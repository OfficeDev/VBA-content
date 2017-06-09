---
title: ContentControlListEntry Object (Word)
keywords: vbawd10.chm2250
f1_keywords:
- vbawd10.chm2250
ms.prod: word
api_name:
- Word.ContentControlListEntry
ms.assetid: b4e51492-4283-22e7-0f9a-2cfa1abaa306
ms.date: 06/08/2017
---


# ContentControlListEntry Object (Word)

A  **ContentControlListEntry** object represents a list item in a drop-down list or combo box content control. A **ContentControlListEntry** object is a member of the **[ContentControlListEntries](contentcontrollistentries-object-word.md)** collection for a **ContentControl** object.


## Remarks

Use the  **[Add](contentcontrollistentries-add-method-word.md)** method of the **ContentControlListEntries** collection to create a new **ContentControlListEntry** object. Use the **[Item](http://msdn.microsoft.com/library/8ecc6f40-3ad4-9647-26ec-fe6ca0fe7953%28Office.15%29.aspx)** method, or **[DropdownListEntries](contentcontrol-dropdownlistentries-property-word.md)** (Index), where Index is the ordinal position of the content control list item, to access an individual list item within the **ContentControlListEntries** collection.


 **Note**  List entries must have unique display names. Attempting to add a list item that already exists raises a run-time error.

The following code example uses the  **Add** method to add several list items to a new drop-down list content control, and then uses the **Item** method to access the third item in the list and change the display text.




```
Dim objCC As ContentControl 
Dim objLE As ContentControlListEntry 
Dim objMap As XMLMapping 
 
Set objCC = ActiveDocument.ContentControls.Add(wdContentControlDropdownList) 
 
'List items 
objCC.DropdownListEntries.Add "Cat" 
objCC.DropdownListEntries.Add "Dog" 
objCC.DropdownListEntries.Add "Equine" 
objCC.DropdownListEntries.Add "Monkey" 
objCC.DropdownListEntries.Add "Snake" 
objCC.DropdownListEntries.Add "Other" 
 
Set objLE = objCC.DropdownListEntries.Item(3) objLE.Text = "Horse" 

```

Use the  **[MoveUp](contentcontrollistentry-moveup-method-word.md)** and **[MoveDown](contentcontrollistentry-movedown-method-word.md)** methods to reposition items in a drop-down list. The following code example moves the first item down, so that it becomes the last item in the list, and moves the last item up, so that it becomes the first item in the list.




```
Dim objcc As ContentControl 
Dim objLE1 As ContentControlListEntry 
Dim objLE2 As ContentControlListEntry 
Dim intCount As Integer 
 
Set objcc = ActiveDocument.ContentControls.Item(3) 
 
If objcc.Type = wdContentControlComboBox Or _ 
 objcc.Type = wdContentControlDropdownList Then 
 
 'First item in the list. 
 Set objLE1 = objcc.DropdownListEntries.Item(1) 
 
 'Last item in the list. 
 Set objLE2 = objcc.DropdownListEntries.Item(objcc.DropdownListEntries.Count) 
 
 For intCount = 1 To objcc.DropdownListEntries.Count 
 'Move the first item down one. 
 objLE1.MoveDown 
 
 'Move the last item up one. 
 objLE2.MoveUp 
 Next 
 
End If
```

Use the  **[Select](contentcontrollistentry-select-method-word.md)** method to programmatically select a content control list item. The following code example inserts a drop-down list content control into the active document, sets the title and placeholder text and adds several items to the list, and then selects the last item entered.




```
Dim objCC As ContentControl 
Dim objCE As ContentControlListEntry 
Dim objMap As XMLMapping 
 
Set objCC = ActiveDocument.ContentControls.Add(wdContentControlDropdownList) 
objCC.Title = "My Favorite Animal" 
If objCC.ShowingPlaceholderText Then _ 
 objCC.SetPlaceholderText , , "Select your favorite animal " 
 
'List entries 
objCC.DropdownListEntries.Add "Cat" 
objCC.DropdownListEntries.Add "Dog" 
objCC.DropdownListEntries.Add "Horse" 
objCC.DropdownListEntries.Add "Monkey" 
objCC.DropdownListEntries.Add "Snake" 
 
Set objCE = objCC.DropdownListEntries.Add("Other") 
objCE.Select
```

Use the  **[Text](contentcontrollistentry-text-property-word.md)** property to set the display text for a content control list item, and use the **[Value](contentcontrollistentry-value-property-word.md)** property to set a programmatic value that you may use later for processing a form. For example, you may use a content control drop-down list for a list of products. The **Text** property may contain the name of the product, a display name that a user can easily recognize and understand. The **Value** property may contain the product number for the product that corresponds to a product number in a database. You can then use the product number from the **Value** property to look up product information in a database. Also, the value of the **Value** property is what is sent to the custom XML data if the content control is mapped to XML data in the data store.

The following code example sets the value for the item based on the contents of the display text.




```
Dim objCc As ContentControl 
Dim objLe As ContentControlListEntry 
Dim strText As String 
Dim strChar As String 
 
Set objCc = ActiveDocument.ContentControls(3) 
 
For Each objLE In objCC.DropdownListEntries 
 If objLE.Text <> "Other" Then 
 strText = objLE.Text 
 objLE.Value = "My favorite animal is the " &amp; strText &amp; "." 
 End If 
Next
```

Use the  **[Delete](contentcontrollistentry-delete-method-word.md)** method to remove an item from a content control drop-down list or combo box. The following code example deletes a drop-down list item if the display text of the item is "Other".




```
Dim objCC As ContentControl 
Dim objCL As ContentControlListEntry 
 
For Each objCC In ActiveDocument.ContentControls 
 If objCC.Type = wdContentControlComboBox Or _ 
 objCC.Type = wdContentControlDropdownList Then 
 For Each objCL In objCC.DropdownListEntries 
 If objCL.Text = "Other" Then objCL.Delete 
 Next 
 End If 
Next 
 
```


## Methods



|**Name**|
|:-----|
|[Delete](contentcontrollistentry-delete-method-word.md)|
|[MoveDown](contentcontrollistentry-movedown-method-word.md)|
|[MoveUp](contentcontrollistentry-moveup-method-word.md)|
|[Select](contentcontrollistentry-select-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Application](contentcontrollistentry-application-property-word.md)|
|[Creator](contentcontrollistentry-creator-property-word.md)|
|[Index](contentcontrollistentry-index-property-word.md)|
|[Parent](contentcontrollistentry-parent-property-word.md)|
|[Text](contentcontrollistentry-text-property-word.md)|
|[Value](contentcontrollistentry-value-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
