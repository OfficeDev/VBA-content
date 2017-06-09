---
title: Document.FollowHyperlink Method (Word)
keywords: vbawd10.chm158007431
f1_keywords:
- vbawd10.chm158007431
ms.prod: word
api_name:
- Word.Document.FollowHyperlink
ms.assetid: ef9a3993-a7b5-5668-e804-c9d1f4fdb7dd
ms.date: 06/08/2017
---


# Document.FollowHyperlink Method (Word)

Displays a cached document, if it has already been downloaded. Otherwise, this method resolves the hyperlink, downloads the target document, and displays the document in the appropriate application.


## Syntax

 _expression_ . **FollowHyperlink**( **_Address_** , **_SubAddress_** , **_NewWindow_** , **_AddHistory_** , **_ExtraInfo_** , **_Method_** , **_HeaderInfo_** )

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Address_|Required| **String**|The address of the target document.|
| _SubAddress_|Optional| **Variant**|The location within the target document. The default value is an empty string.|
| _NewWindow_|Optional| **Variant**| **True** to display the target location in a new window. The default value is **False** .|
| _AddHistory_|Optional| **Variant**| **True** to add the link to the current day's history folder.|
| _ExtraInfo_|Optional| **Variant**|A string or a byte array that specifies additional information for HTTP to use to resolve the hyperlink. For example, you can use ExtraInfo to specify the coordinates of an image map, the contents of a form, or a FAT file name. The string is either posted or appended, depending on the value of Method. Use the  **ExtraInfoRequired** property to determine whether extra information is required.|
| _Method_|Optional| **Variant**|Specifies the way additional information for HTTP is handled. Can be any  **MsoExtraInfoMethod** constant.|
| _HeaderInfo_|Optional| **Variant**|A string that specifies header information for the HTTP request. The default value is an empty string.You can combine several header lines into a single string by using the following syntax: "string1" &; vbCr &; "string2". The specified string is automatically converted into ANSI characters. Note that the HeaderInfo argument may overwrite default HTTP header fields.|

## Example

This example follows the specified URL address and displays the Microsoft home page in a new window.


```vb
ActiveDocument.FollowHyperlink _ 
 Address:="http://www.Microsoft.com", _ 
 NewWindow:=True, AddHistory:=True
```

This example displays the HTML document named "Default.htm."




```vb
ActiveDocument.FollowHyperlink Address:="file:C:\Pages\Default.htm"
```


## See also


#### Concepts


[Document Object](document-object-word.md)

