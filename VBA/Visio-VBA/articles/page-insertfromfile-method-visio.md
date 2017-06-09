---
title: Page.InsertFromFile Method (Visio)
keywords: vis_sdr.chm10916365
f1_keywords:
- vis_sdr.chm10916365
ms.prod: visio
api_name:
- Visio.Page.InsertFromFile
ms.assetid: 03762511-9f2f-6691-ac82-dcff74fcde1d
ms.date: 06/08/2017
---


# Page.InsertFromFile Method (Visio)

Adds a linked or embedded object to a page, master, or group.


## Syntax

 _expression_ . **InsertFromFile**( **_FileName_** , **_Flags_** )

 _expression_ A variable that represents a **Page** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The name of the file that contains the object to link or embed.|
| _Flags_|Required| **Integer**|Flags that influence how the object is inserted.|

### Return Value

Shape


## Remarks

The  **InsertFromFile** method creates a new shape that represents a linked or embedded OLE object.

The  _Flags_ argument is a bitmask that should be a combination of the following values.



|**Constant **|**Value **|**Description **|
|:-----|:-----|:-----|
| **visInsertLink**|&;H8|If set, the new shape represents an OLE link to the named file. Otherwise, the  **InsertFromFile** method produces an OLE object from the contents of the named file and embeds it in the document that contains the page, master, or group.|
| **visInsertIcon**|&;H10|Displays the new shape as an icon.|
 **Security** Use caution when you are adding ActiveX controls to your application. ActiveX controls may be designed in such a way that their use could pose a security risk. We recommend that you use controls from trusted sources only.


