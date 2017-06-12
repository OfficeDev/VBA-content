---
title: ContentControl.Type Property (Word)
keywords: vbawd10.chm266534917
f1_keywords:
- vbawd10.chm266534917
ms.prod: word
api_name:
- Word.ContentControl.Type
ms.assetid: 24f4099d-b4ad-c7be-60a4-e23ede378208
ms.date: 06/08/2017
---


# ContentControl.Type Property (Word)

Returns or sets a  **[WdContentControlType](wdcontentcontroltype-enumeration-word.md)** that represents the type for a content control. Read/write.


## Syntax

 _expression_ . **Type**

 _expression_ An expression that returns a **ContentControl** object.


## Remarks

You can use the  **Type** property to change the type of a content control from one type to another. However, the ability to change the type of control depends on the original type and on the content inside the content control at the time of the change. All content controls can be changed to rich text or building block gallery type content controls because these types allow arbitrary content. For other types, if the content is valid for the type that you want to change to, then changing the type is allowed. Otherwise, the change is rejected, resulting in a run-time error.


## Example

The following example checks to see if the specified content control is a drop-down list box or a combo box, and if it is one of these two types, moves the last item in the list up, so that it becomes the first item in the list.


```vb
Dim objCC As ContentControl 
Dim objCL As ContentControlListEntry 
Dim intCount As Integer 
 
Set objCC = ActiveDocument.ContentControls.Item(3) 
 
If objCC.Type = wdContentControlComboBox Or _ 
 objCC.Type = wdContentControlDropdownList Then 
 
 Set objCL = objCC.DropdownListEntries.Item(objCC.DropdownListEntries.Count) 
 
 For intCount = 1 To objCC.DropdownListEntries.Count 
 objCL.MoveUp 
 Next 
 
End If
```


## See also


#### Concepts


[ContentControl Object](contentcontrol-object-word.md)

