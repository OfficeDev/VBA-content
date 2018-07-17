---
title: Global.WordBasic Property (Word)
keywords: vbawd10.chm163119110
f1_keywords:
- vbawd10.chm163119110
ms.prod: word
api_name:
- Word.Global.WordBasic
ms.assetid: be6209eb-d06c-3399-23b2-31b62642fe83
ms.date: 06/08/2017
---


# Global.WordBasic Property (Word)

Returns an Automation object (Word.Basic) that includes methods for all the WordBasic statements and functions available in Word version 6.0 and Word for Windows 95. Read-only.


## Syntax

 _expression_ . **WordBasic**

 _expression_ A variable that represents a **[Global](global-object-word.md)** object.


## Remarks

In Word 2000 and later, when you open a Word version 6.0 or Word for Windows 95 template that contains WordBasic macros, the macros are automatically converted to Visual Basic modules. Each WordBasic statement and function in the macro is converted to the corresponding Word.Basic method.

For information about WordBasic statements and functions, see WordBasic Help in Word version 6.0 or Word for Windows 95. For information about converting WordBasic to Visual Basic, see [Converting WordBasic Macros to Visual Basic](http://msdn.microsoft.com/library/44a08969-f0e9-291e-7663-b7cc2e3660db%28Office.15%29.aspx). For general information, see [Conceptual Differences Between WordBasic and Visual Basic](http://msdn.microsoft.com/library/2ec0fa57-68c4-f4e9-000c-91a2b97ac9ac%28Office.15%29.aspx).


## Example

This example uses the Word.Basic object to create a new document in Word version 6.0 or Word for Windows 95 and insert the available font names. Each font name is formatted in its corresponding font.


```vb
With WordBasic 
 .FileNewDefault 
 For aCount = 1 To .CountFonts() 
 .Font .[Font$](aCount) 
 .Insert .[Font$](aCount) 
 .InsertPara 
 Next 
End With
```


## See also


#### Concepts


[Global Object](global-object-word.md)

