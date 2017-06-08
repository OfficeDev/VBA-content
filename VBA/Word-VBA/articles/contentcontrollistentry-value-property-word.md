---
title: ContentControlListEntry.Value Property (Word)
keywords: vbawd10.chm147456104
f1_keywords:
- vbawd10.chm147456104
ms.prod: word
api_name:
- Word.ContentControlListEntry.Value
ms.assetid: b37925d7-00ce-9c66-d5d3-bec840d0a2e8
ms.date: 06/08/2017
---


# ContentControlListEntry.Value Property (Word)

Returns or sets a  **String** that represents the programmatic value of an item in a drop-down list or combo box content control. Read/write.


## Syntax

 _expression_ . **Value**

 _expression_ An expression that returns a **ContentControlListEntry** object.


## Remarks

Use the  **Value** property to store data that you need to use at processing time. For example, the **Text** property may contain a string that you want to display and the **Value** property may contain a number, such as an item number, that you can use to look up information in a database. Also, the value of the **Value** property is what is sent to the custom XML data, if the content control is mapped to XML data in the data store.


 **Note**  You cannot set the  **Value** property for list entries that were automatically populated from an XML schema attached to the custom XML that is mapped to this control.


## Example

The following example sets the value for the item based on the contents of the display text.


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

