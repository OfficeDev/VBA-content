---
title: PictureFormat.Application Property (Excel)
ms.prod: excel
api_name:
- Excel.PictureFormat.Application
ms.assetid: afc9ab72-cf23-a4de-1c21-4d4e28bd623b
ms.date: 06/08/2017
---


# PictureFormat.Application Property (Excel)

When used without an object qualifier, this property returns an  **[Application](application-object-excel.md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents a **PictureFormat** object.


## Example

This example displays a message about the application that created  `myObject`.


```vb
Set myObject = ActiveWorkbook 
If myObject.Application.Value = "Microsoft Excel" Then 
 MsgBox "This is an Excel Application object." 
Else 
 MsgBox "This is not an Excel Application object." 
End If
```


## See also


#### Concepts


[PictureFormat Object](pictureformat-object-excel.md)

