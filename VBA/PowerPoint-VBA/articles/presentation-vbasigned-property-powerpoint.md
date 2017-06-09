---
title: Presentation.VBASigned Property (PowerPoint)
keywords: vbapp10.chm583059
f1_keywords:
- vbapp10.chm583059
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.VBASigned
ms.assetid: eebb411d-6312-f858-275f-b0f0ee12b212
ms.date: 06/08/2017
---


# Presentation.VBASigned Property (PowerPoint)

Determines whether the Visual Basic for Applications (VBA) project for the specified document has been digitally signed. Read-only.


## Syntax

 _expression_. **VBASigned**

 _expression_ A variable that represents a **Presentation** object.


### Return Value

MsoTriState


## Remarks

The value of the  **VBASigned** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The VBA project for the specified document has not been digitally signed.|
|**msoTrue**| The VBA project for the specified document has been digitally signed.|

## Example

This example loads a presentation called MyPres.ppt and tests to see whether or not it has a digital signature. If there's no digital signature, the code displays a warning message.


```
Presentations.Open FileName:="c:\My Documents\MyPres.ppt", _
    ReadOnly:=msoFalse, WithWindow:=msoTrue

With ActivePresentation
    If .VBASigned = msoFalse And _
           .VBProject.VBComponents.Count > 0 Then
       MsgBox "Warning! The Visual Basic project for" _
           &; vbCrLf &; "this presentation has not" _
           &; vbCrLf &; " been digitally signed." _
           , vbCritical, "Digital Signature Warning"
    End If
End With
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

