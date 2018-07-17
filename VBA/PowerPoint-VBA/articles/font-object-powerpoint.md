---
title: Font Object (PowerPoint)
keywords: vbapp10.chm575000
f1_keywords:
- vbapp10.chm575000
ms.prod: powerpoint
api_name:
- PowerPoint.Font
ms.assetid: ad62daaa-01a5-36cc-5451-e0da0134ac95
ms.date: 06/08/2017
---


# Font Object (PowerPoint)

Represents character formatting for text or a bullet. The  **Font** object is a member of the **[Fonts](http://msdn.microsoft.com/library/1a8f44ea-515f-5eb9-eab5-6204d9b1d5bc%28Office.15%29.aspx)** collection. The **Fonts** collection contains all the fonts used in a presentation.


## Example

The following examples describes how to do the following:


- Return the  **Font** object that represents the font attributes of a specified bullet, a specified range of text, or all text at a specified outline level
    
- Return a  **Font** object from the collection of all the fonts used in the presentation
    
Use the [Font](http://msdn.microsoft.com/library/234c8843-3c0d-a425-0173-02c3910ba400%28Office.15%29.aspx)property to return the  **Font** object that represents the font attributes for a specific bullet, text range, or outline level. The following example sets the title text on slide one and sets the font properties.




```
With ActivePresentation.Slides(1).Shapes.Title _

        .TextFrame.TextRange

    .Text = "Volcano Coffee"

    With .Font

        .Italic = True

        .Name = "Palatino"

        .Color.RGB = RGB(0, 0, 255)

    End With

End With
```

Use  **Fonts** (index), where index is the font's name or index number, to return a single **Font** object. The following example checks to see whether font one in the active presentation is embedded in the presentation.




```
If ActivePresentation.Fonts(1).Embedded = _

    True Then MsgBox "Font 1 is embedded"
```


## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/70e38091-9f12-74c7-18b9-13474ac26644%28Office.15%29.aspx)|
|[AutoRotateNumbers](http://msdn.microsoft.com/library/621ccc86-d5cb-d2c1-262f-5652eff5800a%28Office.15%29.aspx)|
|[BaselineOffset](http://msdn.microsoft.com/library/aa948e2e-957c-ff4c-16b9-480d7f5f2d24%28Office.15%29.aspx)|
|[Bold](http://msdn.microsoft.com/library/13e81c46-5ae7-21ee-58e1-5ab23de552d5%28Office.15%29.aspx)|
|[Color](http://msdn.microsoft.com/library/461d3fdc-5097-ceca-76f6-81d924f8a7b7%28Office.15%29.aspx)|
|[Embeddable](http://msdn.microsoft.com/library/50824587-0371-e7eb-8885-370f97b8bf0c%28Office.15%29.aspx)|
|[Embedded](http://msdn.microsoft.com/library/3fd7fe50-19a9-9944-f7c8-0ba54bc07c93%28Office.15%29.aspx)|
|[Emboss](http://msdn.microsoft.com/library/734b5bd7-4b1f-d3b3-d8bd-f73d0bc86f67%28Office.15%29.aspx)|
|[Italic](http://msdn.microsoft.com/library/5fc7e3fe-e103-72ea-42cb-c178b411312a%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/6798b75b-7fb8-a046-1532-a8cc41b76af8%28Office.15%29.aspx)|
|[NameAscii](http://msdn.microsoft.com/library/06db0f5b-71ac-704d-eef2-1be8a96fb7a8%28Office.15%29.aspx)|
|[NameComplexScript](http://msdn.microsoft.com/library/ef1e44d6-ff5d-aaa9-4eaa-643cb2ebc2bf%28Office.15%29.aspx)|
|[NameFarEast](http://msdn.microsoft.com/library/0b3f7d98-bda5-eec3-f570-20d8b575c0a3%28Office.15%29.aspx)|
|[NameOther](http://msdn.microsoft.com/library/64f62838-635c-9b6d-082a-06fe698685e1%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/5cf96dc7-aa6a-e3f6-d8f3-c0b92d6b1a6a%28Office.15%29.aspx)|
|[Shadow](http://msdn.microsoft.com/library/37d23e3a-26a7-ba20-1e23-13861090ae79%28Office.15%29.aspx)|
|[Size](http://msdn.microsoft.com/library/dd56a4e9-20c7-b38d-0d0e-82e5326d51c4%28Office.15%29.aspx)|
|[Subscript](http://msdn.microsoft.com/library/ad23433b-b14b-9b2a-3bf6-772de41995f7%28Office.15%29.aspx)|
|[Superscript](http://msdn.microsoft.com/library/6f0bba73-f375-d715-3ddb-f1ab6041336c%28Office.15%29.aspx)|
|[Underline](http://msdn.microsoft.com/library/ee21ab18-b131-7e4d-de19-93c9b7549d3b%28Office.15%29.aspx)|

## See also


#### Other resources


[PowerPoint Object Model Reference](http://msdn.microsoft.com/library/00acd64a-5896-0459-39af-98df2849849e%28Office.15%29.aspx)
