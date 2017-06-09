---
title: ContentControlListEntry.Text Property (Word)
keywords: vbawd10.chm147456000
f1_keywords:
- vbawd10.chm147456000
ms.prod: word
api_name:
- Word.ContentControlListEntry.Text
ms.assetid: bfe2487b-7ba6-3047-842b-0c2466919efb
ms.date: 06/08/2017
---


# ContentControlListEntry.Text Property (Word)

Returns or sets a  **String** that represents the display text of a list item for a drop-down list or combo box content control. Read/write.


## Syntax

 _expression_ . **Text**

 _expression_ An expression that returns a **ContentControlListEntry** object.


## Remarks

List entries must have unique display names. Attempting to change the  **Text** property to a string that already exists in the list of entries raises a run-time error.


## Example

The following example capitalizes the first character, if it is lowercase, in the display text of each list item.


```vb
Dim objCC As ContentControl 
Dim objLE As ContentControlListEntry 
Dim strFirst As String 
 
For Each objCC In ActiveDocument.ContentControls 
 If objCC.Type = wdContentControlComboBox Or objCC.Type = wdContentControlDropdownList Then 
 For Each objLE In objCC.DropdownListEntries 
 strFirst = Left(objLE.Text, 1) 
 
 If strFirst = LCase(strFirst) Then 
 objLE.Text = UCase(strFirst) &; Right(objLe.Text, Len(objLe.Text) - 1) 
 End If 
 Next 
 End If 
Next
```

The following example sets the value for the list item based on the contents of the display text.




```vb
Dim objCc As ContentControl 
Dim objLe As ContentControlListEntry 
Dim strText As String 
Dim strChar As String 
 
Set objCc = ActiveDocument.ContentControls(3) 
 
For Each objLE In objCC.DropdownListEntries 
 If objLE.Text <> "Other" Then 
 strText = objLE.Text 
 objLE.Value = "My favorite animal is the " &; strText &; "." 
 End If 
Next
```


## See also


#### Concepts


[ContentControlListEntry Object](contentcontrollistentry-object-word.md)

