---
title: ErrorCheckingOptions.NumberAsText Property (Excel)
keywords: vbaxl10.chm698077
f1_keywords:
- vbaxl10.chm698077
ms.prod: excel
api_name:
- Excel.ErrorCheckingOptions.NumberAsText
ms.assetid: 4ef057e1-82da-78ab-0541-a91fbdef4d89
ms.date: 06/08/2017
---


# ErrorCheckingOptions.NumberAsText Property (Excel)

When set to  **True** (default), Microsoft Excel identifies, with an **AutoCorrect Options** button, selected cells that contain numbers written as text. **False** disables error checking for numbers written as text. Read/write **Boolean** .


## Syntax

 _expression_ . **NumberAsText**

 _expression_ A variable that represents an **ErrorCheckingOptions** object.


## Example

In the following example, the  **AutoCorrect Options** button appears for cell A1, which contains a number stored as text.


```vb
Sub CheckNumberAsText() 
 
 ' Simulate an error by referencing a number stored as text. 
 Application.ErrorCheckingOptions.NumberAsText = True 
 Range("A1").Value = "'1" 
 
End Sub
```


## See also


#### Concepts


[ErrorCheckingOptions Object](errorcheckingoptions-object-excel.md)

