---
title: Presentation.FollowHyperlink Method (PowerPoint)
keywords: vbapp10.chm583030
f1_keywords:
- vbapp10.chm583030
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.FollowHyperlink
ms.assetid: 411863be-0bd9-c939-1309-9f537b47f30b
ms.date: 06/08/2017
---


# Presentation.FollowHyperlink Method (PowerPoint)

Displays a cached document, if it has already been downloaded. Otherwise, this method resolves the hyperlink, downloads the target document and displays it in the appropriate application.


## Syntax

 _expression_. **FollowHyperlink**( **_Address_**, **_SubAddress_**, **_NewWindow_**, **_AddHistory_**, **_ExtraInfo_**, **_Method_**, **_HeaderInfo_** )

 _expression_ A variable that represents a **Presentation** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Address_|Required|**String**|The address of the target document.|
| _SubAddress_|Optional|**String**|The location in the target document. By default, this argument is an empty string.|
| _NewWindow_|Optional|**Boolean**|**True** to have the target application opened in a new window. The default value is **False**.|
| _AddHistory_|Optional|**Boolean**|**True** to add the link to the current day's history folder.|
| _ExtraInfo_|Optional|**String**|String or byte array that specifies information for HTTP. This argument can be used, for example, to specify the coordinates of an image map or the contents of a form. It can also indicate a FAT file name. The Method argument determines how this extra information is handled.|
| _Method_|Optional|**MsoExtraInfoMethod**|Specifies how ExtraInfo is posted or appended.|
| _HeaderInfo_|Optional|**String**| A string that specifies header information for the HTTP request. The default value is an empty string. You can combine several header lines into a single string by using the following syntax: "string1" &; vbCr &; "string2". The specified string is automatically converted into ANSI characters. Note that the HeaderInfo argument may overwrite default HTTP header fields.|

### Return Value

Nothing


## Remarks

ExtraInfo can be one of these  **MsoExtraInfoMethod** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoMethodGet**|The default. ExtraInfo is a  **String** that is appended to the address.|
|**msoMethodPost**|ExtraInfo is posted as a  **String** or byte array.|

## Example

This example loads the document at example.microsoft.com in a new window and adds it to the history folder.


```vb
ActivePresentation.FollowHyperlink _
    Address:="http://example.microsoft.com", _
    NewWindow:=True, AddHistory:=True
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

