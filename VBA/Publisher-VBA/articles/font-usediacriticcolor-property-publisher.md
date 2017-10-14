---
title: Font.UseDiacriticColor Property (Publisher)
keywords: vbapb10.chm5374002
f1_keywords:
- vbapb10.chm5374002
ms.prod: publisher
api_name:
- Publisher.Font.UseDiacriticColor
ms.assetid: 368d3599-b0b0-1790-0ce0-13f1936bccb0
ms.date: 06/08/2017
---


# Font.UseDiacriticColor Property (Publisher)

Returns or sets  **MsoTriState** constant indicating whether you can set the color of diacritics in the specified text range. Read/write.


## Syntax

 _expression_. **UseDiacriticColor**

 _expression_A variable that represents an  **Font** object.


### Return Value

MsoTriState


## Remarks

The  **UseDiacriticColor** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**|The color of diacritics cannot be set in the specified text range.|
| **msoTriStateMixed**|Return value indicating a combination of  **msoTrue** and **msoFalse** for the specified shape range.|
| **msoTriStateToggle**|Set value that switches between **msoTrue** and **msoFalse**.|
| **msoTrue**|The color of diacritics can be set in the specified text range.|

## Example

This example test the text in the first story of the publication for the state of the  **UseDiacriticColor** property. If it is **msoTrue**, the  **DiacriticColor** property value is set to blue. Otherwise, a message box is displayed.


```vb
Sub UseDiaColor() 
 
 Dim fntDC As Font 
 
 Set fntDC = Application.ActiveDocument.Stories(1).TextRange.Font 
 If fntDC.UseDiacriticColor = msoTrue Then 
 fntDC.DiacriticColor.RGB = RGB(Red:=0, Green:=0, Blue:=255) 
 Else 
 MsgBox "The UseDiacriticColor property is set to False" 
 End If 
 
End Sub
```


