---
title: TextRange.InsertSymbol Method (PowerPoint)
keywords: vbapp10.chm569022
f1_keywords:
- vbapp10.chm569022
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.InsertSymbol
ms.assetid: a424e011-1bfe-f690-cbc0-604f89718831
ms.date: 06/08/2017
---


# TextRange.InsertSymbol Method (PowerPoint)

Returns a  **[TextRange](textrange-object-powerpoint.md)** object that represents a symbol inserted into the specified text range.


## Syntax

 _expression_. **InsertSymbol**( **_FontName_**, **_CharNumber_**, **_UniCode_** )

 _expression_ A variable that represents an **TextRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FontName_|Required|**String**|The font name.|
| _CharNumber_|Required|**Long**|The Unicode or ASCII character number.|
| _Unicode_|Optional|**MsoTriState**|Specifies whether the CharNumber argument represents an ASCII or Unicode character.|

### Return Value

TextRange


## Remarks

The  _CharNumber_ parameter value can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The default. The CharNumber argument represents an ASCII character number.|
|**msoTrue**|The CharNumber argument represents a Unicode character.|

## Example

This example inserts the registered trademark symbol after the first sentence of the first paragraph in a new text box on the first slide in the active presentation.


```vb
Sub Symbol()

    Dim txtBox As Shape

    'Add text box
    Set txtBox = Application.ActivePresentation.Slides(1) _
        .Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
            Left:=100, Top:=100, Width:=100, Height:=100)

    'Add symbol to text box
    txtBox.TextFrame.TextRange.InsertSymbol _
        FontName:="Symbol", CharNumber:=226

End Sub
```


## See also


#### Concepts


[TextRange Object](textrange-object-powerpoint.md)

