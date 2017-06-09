---
title: OLEFormat.ConvertTo Method (Word)
keywords: vbawd10.chm154337390
f1_keywords:
- vbawd10.chm154337390
ms.prod: word
api_name:
- Word.OLEFormat.ConvertTo
ms.assetid: 6d648f38-34fa-21b1-3ab9-a1965f2398f4
ms.date: 06/08/2017
---


# OLEFormat.ConvertTo Method (Word)

Converts the specified OLE object from one class to another, making it possible for you to edit the object in a different server application, or changing how the object is displayed in the document.


## Syntax

 _expression_ . **ConvertTo**( **_ClassType_** , **_DisplayAsIcon_** , **_IconFileName_** , **_IconIndex_** , **_IconLabel_** )

 _expression_ Required. A variable that represents an **[OLEFormat](oleformat-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ClassType_|Optional| **Variant**|The name of the application used to activate the OLE object. You can see a list of the available applications in the  **Object type** box on the **Create New** tab in the **Object** dialog box. You can find the ClassType string by inserting an object as an inline shape and then viewing the field codes. The class type of the object follows either the word "EMBED" or the word "LINK."|
| _DisplayAsIcon_|Optional| **Variant**| **True** to display the OLE object as an icon. The default value is **False** .|
| _IconFileName_|Optional| **Variant**|The file that contains the icon to be displayed.|
| _IconIndex_|Optional| **Variant**|The index number of the icon within IconFileName. The order of icons in the specified file corresponds to the order in which the icons appear in the  **Change Icon** dialog box ( **Insert Object** dialog box) when the **Display as icon** check box is selected. The first icon in the file has the index number 0 (zero). If an icon with the given index number doesn't exist in IconFileName, the icon with the index number 1 (the second icon in the file) is used. The default value is 0 (zero).|
| _IconLabel_|Optional| **Variant**|A label (caption) to be displayed beneath the icon.|

## Example

This example creates a new document, then inserts an embedded Word document with some text. Then, the embedded document is converted to a Word Picture.


```vb
Dim objEmbedded As Object 
 
Documents.Add 
 
Set objEmbedded = ActiveDocument.Shapes _ 
 .AddOLEObject(ClassType:= "Word.Document") 
objEmbedded.Activate 
Selection.TypeText "Test" 
objEmbedded.OLEFormat.OLEFormat.ConvertTo _ 
 ClassType:="Word.Picture"
```


## See also


#### Concepts


[OLEFormat Object](oleformat-object-word.md)

