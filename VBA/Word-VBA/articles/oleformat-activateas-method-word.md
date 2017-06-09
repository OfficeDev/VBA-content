---
title: OLEFormat.ActivateAs Method (Word)
keywords: vbawd10.chm154337391
f1_keywords:
- vbawd10.chm154337391
ms.prod: word
api_name:
- Word.OLEFormat.ActivateAs
ms.assetid: 3db19832-efcf-c392-4e76-82ec297a3d69
ms.date: 06/08/2017
---


# OLEFormat.ActivateAs Method (Word)

Sets the Windows registry value that determines the default application used to activate the specified OLE object.


## Syntax

 _expression_ . **ActivateAs**( **_ClassType_** )

 _expression_ Required. A variable that represents an **[OLEFormat](oleformat-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ClassType_|Required| **String**|The name of the application in which an OLE object is opened. To see a list of object types that the OLE object can be activated as, click the object and then open the  **Convert** dialog box. You can find the ClassType string by inserting an object as an inline shape and then viewing the field codes. The class type of the object follows either the word "EMBED" or the word "LINK."|

## Example

This example sets the first floating shape on the active document to open in Microsoft Excel, and then it activates the shape. For the example to work, this shape must be an OLE object that can be opened in Microsoft Excel.


```vb
With ActiveDocument.Shapes(1).OLEFormat 
 .ActivateAs ClassType:="Excel.Sheet" 
 .Activate 
End With
```


## See also


#### Concepts


[OLEFormat Object](oleformat-object-word.md)

