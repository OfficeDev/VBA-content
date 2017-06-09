---
title: Options.DisplayPasteOptions Property (PowerPoint)
keywords: vbapp10.chm667001
f1_keywords:
- vbapp10.chm667001
ms.prod: powerpoint
api_name:
- PowerPoint.Options.DisplayPasteOptions
ms.assetid: 4c5f0851-585c-e4c6-a6c7-c3bfd3666883
ms.date: 06/08/2017
---


# Options.DisplayPasteOptions Property (PowerPoint)

Determines whether Microsoft PowerPoint displays the  **Paste Options** button, which appears directly under newly pasted text. Read/write.


## Syntax

 _expression_. **DisplayPasteOptions**

 _expression_ A variable that represents a **Options** object.


### Return Value

MsoTriState


## Remarks

The value of the  **DisplayPasteOptions** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The  **PasteOptions** button is not displayed.|
|**msoTrue**| The **PasteOptions** button is displayed.|

## Example

This example enables the  **Paste Options** button if the option has been disabled.


```vb
Sub ShowPasteOptionsButton()

    With Application.Options

        If  .DisplayPasteOptions = False Then

            .DisplayPasteOptions = True

        End If

    End With

End Sub
```


## See also


#### Concepts


[Options Object](options-object-powerpoint.md)

