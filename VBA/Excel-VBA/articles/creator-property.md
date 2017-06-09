---
title: Creator Property
keywords: vbagr10.chm65685
f1_keywords:
- vbagr10.chm65685
ms.prod: excel
api_name:
- Excel.Creator
ms.assetid: 79d72908-f141-1d3a-d8db-c10db7b33537
ms.date: 06/08/2017
---


# Creator Property

Returns a 32-bit integer that indicates the application in which the specified object was created. If the object was created in Microsoft Graph, this property returns the string MSGR, which is equivalent to the hexadecimal number 4D534752. Read-only XlCreator.



|XlCreator can be one of these XlCreator constants.|
| **xlCreatorCode**|

 _expression_. **Creator**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

## Example

This example displays a message about the creator of  `myChart`.


```vb
If myChart.Creator = &;h4D534752 Then 
    MsgBox "This is a Microsoft Graph object" 
Else 
    MsgBox "This is not a Microsoft Graph object" 
End If
```


