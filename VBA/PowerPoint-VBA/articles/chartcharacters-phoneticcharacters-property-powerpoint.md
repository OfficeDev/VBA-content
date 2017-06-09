---
title: ChartCharacters.PhoneticCharacters Property (PowerPoint)
keywords: vbapp10.chm67058
f1_keywords:
- vbapp10.chm67058
ms.prod: powerpoint
api_name:
- PowerPoint.ChartCharacters.PhoneticCharacters
ms.assetid: b3ceaf21-db47-7fd3-4414-3fc3040a55b4
ms.date: 06/08/2017
---


# ChartCharacters.PhoneticCharacters Property (PowerPoint)

Returns or sets the phonetic text for the object. Read/write  **String**.


## Syntax

 _expression_. **PhoneticCharacters**

 _expression_ A variable that represents a **[ChartCharacters](chartcharacters-object-powerpoint.md)** object.


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example replaces the first three characters in the title of the first chart in the active document with Furigana characters.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.Title.Characters(1,3).PhoneticCharacters = "Invalid DDUE based on source, error:image not allowed in code, image filename:543934d2-a0ba-508d-09a6-f71880d969e4Invalid DDUE based on source, error:image not allowed in code, image filename:a10b3c7d-b1d8-6602-439a-071c70a35d5bInvalid DDUE based on source, error:image not allowed in code, image filename:add897d6-e820-bf7c-b867-8727538c8534Invalid DDUE based on source, error:image not allowed in code, image filename:6fad4588-0ab9-3701-681a-34f2765b0aa0"

    End If

End With
```


## See also


#### Concepts


[ChartCharacters Object](chartcharacters-object-powerpoint.md)

