---
title: Hyperlinks.Add Method (Word)
keywords: vbawd10.chm161218661
f1_keywords:
- vbawd10.chm161218661
ms.prod: word
api_name:
- Word.Hyperlinks.Add
ms.assetid: b838a93c-8ec8-e591-f2e9-c22a049c5335
ms.date: 06/08/2017
---


# Hyperlinks.Add Method (Word)

Returns a  **Hyperlink** object that represents a new hyperlink added to a range, selection, or document.


## Syntax

 _expression_ . **Add**( **_Anchor_** , **_Address_** , **_SubAddress_** , **_ScreenTip_** , **_TextToDisplay_** , **_Target_** )

 _expression_ Required. A variable that represents a **[Hyperlinks](hyperlinks-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Anchor_|Required| **Object**|The text or graphic that you want turned into a hyperlink.|
| _Address_|Optional| **Variant**|The address for the specified link. The address can be an e-mail address, an Internet address, or a file name. Note that Microsoft Word doesn't check the accuracy of the address.|
| _SubAddress_|Optional| **Variant**|The name of a location within the destination file, such as a bookmark, named range, or slide number.|
| _ScreenTip_|Optional| **Variant**|The text that appears as a ScreenTip when the mouse pointer is positioned over the specified hyperlink. The default value is "Address".|
| _TextToDisplay_|Optional| **Variant**|The display text of the specified hyperlink. The value of this argument replaces the text or graphic specified by Anchor.|
| _Target_|Optional| **Variant**|The name of the frame or window in which you want to load the specified hyperlink.|

### Return Value

Hyperlink


## Example

This example turns the selection into a hyperlink to the Microsoft address on the World Wide Web.


```vb
ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, _ 
 Address:="http:\\www.microsoft.com"
```

This example turns the selection into a hyperlink that links to the bookmark named MyBookMark in MyFile.doc.




```vb
ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, _ 
 Address:="C:\My Documents\MyFile.doc", SubAddress:="MyBookMark"
```

This example turns the first shape in the selection into a hyperlink.




```vb
ActiveDocument.Hyperlinks.Add Anchor:=Selection.ShapeRange(1), _ 
 Address:="http:\\www.microsoft.com"
```


## See also


#### Concepts


[Hyperlinks Collection Object](hyperlinks-object-word.md)

