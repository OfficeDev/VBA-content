---
title: ChartCharacters Object (PowerPoint)
keywords: vbapp10.chm686000
f1_keywords:
- vbapp10.chm686000
ms.prod: powerpoint
api_name:
- PowerPoint.ChartCharacters
ms.assetid: 2f659f71-f277-dab4-f2bd-631c7a2424de
ms.date: 06/08/2017
---


# ChartCharacters Object (PowerPoint)

Represents characters in an object that contains text. 


## Remarks

The  **ChartCharacters** object lets you modify any sequence of characters contained in the full text string.

Use  **Characters** ( _Start_, _Length_ ), where _Start_ is the start character number and _Length_ is the number of characters, to return a **ChartCharacters** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The  **[Characters](charttitle-characters-property-powerpoint.md)** property is necessary only when you need to change some of an object's text without affecting the rest (you cannot use the **Characters** property to format a portion of the text if the object does not support rich text). To change all the text at the same time, you can usually apply the appropriate method or property directly to the object. The following example formats the contents of the chart title for the first chart in the active document as italic.




```vb
With ActiveDocument.InlineShapes(1).Chart

    .ChartTitle.Characters.Font.Italic = True

End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

