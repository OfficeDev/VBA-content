---
title: TextFrame2.Creator Property (Excel)
ms.prod: excel
api_name:
- Excel.TextFrame2.Creator
ms.assetid: a6621e71-b864-9e95-68d0-a74649bc15ec
ms.date: 06/08/2017
---


# TextFrame2.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **TextFrame2** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.


## Example

This example displays a message about the creator of an Excel workbook.


```vb
Sub FindCreator() 
 
 Dim myObject As Excel.Workbook 
 Set myObject = ActiveWorkbook 
 If myObject.TextFrame2.Creator = &;h5843454c Then 
 MsgBox "This is a Microsoft Excel object." 
 Else 
 MsgBox "This is not a Microsoft Excel object." 
 End If 
 
End Sub
```


## See also


#### Concepts


[TextFrame2 Object](textframe2-object-excel.md)

