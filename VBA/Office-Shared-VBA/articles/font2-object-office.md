---
title: Font2 Object (Office)
ms.prod: office
api_name:
- Office.Font2
ms.assetid: 8e892c52-56d9-72bd-2893-b15a17cd59ae
ms.date: 06/08/2017
---


# Font2 Object (Office)

Contains font attributes (for example, font name, font size, and color) for an object.


## Example

The following example changes the formatting of the Heading 2 style in the active document to Arial and bold.


```
With ActiveDocument.Styles(wdStyleHeading2).Font2 
 .Name = "Arial" 
 .Italic = True 
End With 

```


## Properties



|**Name**|
|:-----|
|[Allcaps](font2-allcaps-property-office.md)|
|[Application](font2-application-property-office.md)|
|[AutorotateNumbers](font2-autorotatenumbers-property-office.md)|
|[BaselineOffset](font2-baselineoffset-property-office.md)|
|[Bold](font2-bold-property-office.md)|
|[Caps](font2-caps-property-office.md)|
|[Creator](font2-creator-property-office.md)|
|[DoubleStrikeThrough](font2-doublestrikethrough-property-office.md)|
|[Embeddable](font2-embeddable-property-office.md)|
|[Embedded](font2-embedded-property-office.md)|
|[Equalize](font2-equalize-property-office.md)|
|[Fill](font2-fill-property-office.md)|
|[Glow](font2-glow-property-office.md)|
|[Highlight](font2-highlight-property-office.md)|
|[Italic](font2-italic-property-office.md)|
|[Kerning](font2-kerning-property-office.md)|
|[Line](font2-line-property-office.md)|
|[Name](font2-name-property-office.md)|
|[NameAscii](font2-nameascii-property-office.md)|
|[NameComplexScript](font2-namecomplexscript-property-office.md)|
|[NameFarEast](font2-namefareast-property-office.md)|
|[NameOther](font2-nameother-property-office.md)|
|[Parent](font2-parent-property-office.md)|
|[Reflection](font2-reflection-property-office.md)|
|[Shadow](font2-shadow-property-office.md)|
|[Size](font2-size-property-office.md)|
|[Smallcaps](font2-smallcaps-property-office.md)|
|[SoftEdgeFormat](font2-softedgeformat-property-office.md)|
|[Spacing](font2-spacing-property-office.md)|
|[Strike](font2-strike-property-office.md)|
|[StrikeThrough](font2-strikethrough-property-office.md)|
|[Subscript](font2-subscript-property-office.md)|
|[Superscript](font2-superscript-property-office.md)|
|[UnderlineColor](font2-underlinecolor-property-office.md)|
|[UnderlineStyle](font2-underlinestyle-property-office.md)|
|[WordArtformat](font2-wordartformat-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
