---
title: BulletFormat2 Object (Office)
ms.prod: office
api_name:
- Office.BulletFormat2
ms.assetid: ad4c2a05-c34d-fbd4-6b12-3153b94d2c4e
ms.date: 06/08/2017
---


# BulletFormat2 Object (Office)

Represents bullet formatting.


## Example

The following example sets the bullet size and color for the paragraphs in shape two on slide one in the active PowerPoint presentation.


```
With ActivePresentation.Slides(1).Shapes(2) 
 With .TextFrame.TextRange.ParagraphFormat.BulletFormat2 
 .Visible = True 
 .RelativeSize = 1.25 
 .Character = 169 
 With .Font 
 .Color.RGB = RGB(255, 255, 0) 
 .Name = "Symbol" 
 End With 
 End With 
End With 

```


## Methods



|**Name**|
|:-----|
|[Picture](bulletformat2-picture-method-office.md)|

## Properties



|**Name**|
|:-----|
|[Application](bulletformat2-application-property-office.md)|
|[Character](bulletformat2-character-property-office.md)|
|[Creator](bulletformat2-creator-property-office.md)|
|[Font](bulletformat2-font-property-office.md)|
|[Number](bulletformat2-number-property-office.md)|
|[Parent](bulletformat2-parent-property-office.md)|
|[RelativeSize](bulletformat2-relativesize-property-office.md)|
|[StartValue](bulletformat2-startvalue-property-office.md)|
|[Style](bulletformat2-style-property-office.md)|
|[Type](bulletformat2-type-property-office.md)|
|[UseTextColor](bulletformat2-usetextcolor-property-office.md)|
|[UseTextFont](bulletformat2-usetextfont-property-office.md)|
|[Visible](bulletformat2-visible-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
