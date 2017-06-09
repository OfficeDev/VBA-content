---
title: Hyperlink.Follow Method (Word)
keywords: vbawd10.chm161284200
f1_keywords:
- vbawd10.chm161284200
ms.prod: word
api_name:
- Word.Hyperlink.Follow
ms.assetid: ff8553f3-9da7-245f-75fc-77013b5b1e9a
ms.date: 06/08/2017
---


# Hyperlink.Follow Method (Word)

Displays a cached document associated with the specified  **Hyperlink** object, if it has already been downloaded. Otherwise, this method resolves the hyperlink, downloads the target document, and displays the document in the appropriate application.


## Syntax

 _expression_ . **Follow**( **_NewWindow_** , **_AddHistory_** , **_ExtraInfo_** , **_Method_** , **_HeaderInfo_** )

 _expression_ Required. A variable that represents a **[Hyperlink](hyperlink-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _NewWindow_|Optional| **Variant**| **True** to display the target document in a new window. The default value is **False** .|
| _AddHistory_|Optional| **Variant**|This argument is reserved for future use.|
| _ExtraInfo_|Optional| **Variant**|A string or byte array that specifies additional information for HTTP to use to resolve the hyperlink. For example, you can use ExtraInfo to specify the coordinates of an image map, the contents of a form, or a FAT file name. The string is either posted or appended, depending on the value of Method. Use the  **ExtraInfoRequired** property to determine whether extra information is required.|
| _Method_|Optional| **Variant**|Specifies the way additional information for HTTP is handled. Can be any  **MsoExtraInfoMethod** constant.|
| _HeaderInfo_|Optional| **Variant**|A string that specifies header information for the HTTP request. The default value is an empty string. You can combine several header lines into a single string by using the following syntax: "string1" &; vbCr &; "string2". The specified string is automatically converted into ANSI characters. Note that the HeaderInfo argument may overwrite default HTTP header fields.|

## Remarks

If the hyperlink uses the file protocol, this method opens the document instead of downloading it.


## Example

This example follows the first hyperlink in Home.doc.


```vb
Documents("Home.doc").HyperLinks(1).Follow
```

This example inserts a hyperlink to www.msn.com and then follows the hyperlink.




```vb
Dim hypTemp As Hyperlink 
 
With Selection 
 .Collapse Direction:=wdCollapseEnd 
 .InsertAfter "MSN " 
 .Previous 
End With 
 
Set hypTemp = ActiveDocument.Hyperlinks.Add( _ 
 Address:="http://www.msn.com", _ 
 Anchor:=Selection.Range) 
hypTemp.Follow NewWindow:=False, AddHistory:=True
```


## See also


#### Concepts


[Hyperlink Object](hyperlink-object-word.md)

