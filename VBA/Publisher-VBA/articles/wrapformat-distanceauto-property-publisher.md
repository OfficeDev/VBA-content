---
title: WrapFormat.DistanceAuto Property (Publisher)
keywords: vbapb10.chm786437
f1_keywords:
- vbapb10.chm786437
ms.prod: publisher
api_name:
- Publisher.WrapFormat.DistanceAuto
ms.assetid: 8b4e6b93-6e68-5c4a-2164-1a88ca0a633e
ms.date: 06/08/2017
---


# WrapFormat.DistanceAuto Property (Publisher)

Returns or sets an  **MsoTriState** constant indicating whether an appropriate distance between an inline shape and any surrounding text is automatically calculated. Read/write.


## Syntax

 _expression_. **DistanceAuto**

 _expression_A variable that represents a  **WrapFormat** object.


### Return Value

MsoTriState


## Remarks

The  **DistanceAuto** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**|The shape's edges are not adjusted depending on the margins of the text box it overlaps.|
| **msoTriStateMixed**|Return value indicating a combination of  **msoTrue** and **msoFalse** for the specified shape range.|
| **msoTriStateToggle**| Set value that switches the property value between **msoTrue** and **msoFalse**.|
| **msoTrue**|The default. The shape's edges are automatically adjusted depending on the margins of the text box it overlaps. |

## Example

The following example sets shape one on page one of the active publication so that its edges are not automatically adjusted based on its distance from surrounding text.


```vb
Sub SetDistanceAutoProperty() 
 With ActiveDocument.Pages(1).Shapes(1).TextWrap 
 .Type = pbWrapTypeSquare 
 .DistanceAuto = msoFalse 
 End With 
End Sub
```


