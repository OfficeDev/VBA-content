---
title: Application.FormulaBarHeight Property (Excel)
keywords: vbaxl10.chm133306
f1_keywords:
- vbaxl10.chm133306
ms.prod: excel
api_name:
- Excel.Application.FormulaBarHeight
ms.assetid: ff377046-06cb-9cf7-32f5-773da447c184
ms.date: 06/08/2017
---


# Application.FormulaBarHeight Property (Excel)

Allows the user to specify the height of the formula bar in lines. Read/write  **Long** .


## Syntax

 _expression_ . **FormulaBarHeight**

 _expression_ A variable that represents an **Application** object.


## Remarks

If the specified value of  **FormulaBarHeight** is greater than the viewable window space, the formula bar is resized to be equal to the window height.


## Example

In the following example, the height of the formula bar is set to five lines.


```vb
Application.FormulaBarHeight = 5 
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

