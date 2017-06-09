---
title: Characters Object (Excel)
keywords: vbaxl10.chm251072
f1_keywords:
- vbaxl10.chm251072
ms.prod: excel
api_name:
- Excel.Characters
ms.assetid: 128c9ee4-8ba3-6d22-ad0f-9f20be1e24af
ms.date: 06/08/2017
---


# Characters Object (Excel)

Represents characters in an object that contains text. 


## Remarks

The  **Characters** object lets you modify any sequence of characters contained in the full text string.

Use  **Characters** ( _start_, _length_ ), where _start_ is the start character number and _length_ is the number of characters, to return a **Characters** object.


## Example

The following example adds text to cell B1 and then makes the second word bold.


```
With Worksheets("Sheet1").Range("B1") 
 .Value = "New Title" 
 .Characters(5, 5).Font.Bold = True 
End With
```

The  **[Characters](range-characters-property-excel.md)** method is necessary only when you need to change some of an object's text without affecting the rest (you cannot use the **Characters** method to format a portion of the text if the object doesn't support rich text). To change all the text at the same time, you can usually apply the appropriate method or property directly to the object. The following example formats the contents of cell A5 as italic.




```
Worksheets("Sheet1").Range("A5").Font.Italic = True
```


## Methods



|**Name**|
|:-----|
|[Delete](characters-delete-method-excel.md)|
|[Insert](characters-insert-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](characters-application-property-excel.md)|
|[Caption](characters-caption-property-excel.md)|
|[Count](characters-count-property-excel.md)|
|[Creator](characters-creator-property-excel.md)|
|[Font](characters-font-property-excel.md)|
|[Parent](characters-parent-property-excel.md)|
|[PhoneticCharacters](characters-phoneticcharacters-property-excel.md)|
|[Text](characters-text-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
