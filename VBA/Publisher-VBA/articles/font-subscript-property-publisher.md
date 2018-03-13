---
title: Font.SubScript Property (Publisher)
keywords: vbapb10.chm5373973
f1_keywords:
- vbapb10.chm5373973
ms.prod: publisher
api_name:
- Publisher.Font.SubScript
ms.assetid: 9992fdcc-dd60-b2f7-307b-99b10dc7debb
ms.date: 06/08/2017
---


# Font.SubScript Property (Publisher)

Returns or sets an  **MsoTriState** constant indicating whether characters are formatted as subscript in the specified text range. Read/write.


## Syntax

 _expression_. **SubScript**

 _expression_A variable that represents a  **Font** object.


### Return Value

MsoTriState


## Remarks

The  **SubScript** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



| <strong>Constant</strong>          | <strong>Description</strong>                                                                                                    |
|:-----------------------------------|:--------------------------------------------------------------------------------------------------------------------------------|
| <strong>msoFalse</strong>          | No characters in the range are formatted as subscript.                                                                          |
| <strong>msoTriStateMixed</strong>  | Return value indicating a combination of  <strong>msoTrue</strong> and <strong>msoFalse</strong> for the specified shape range. |
| <strong>msoTriStateToggle</strong> | Set value that switches between  <strong>msoTrue</strong> and <strong>msoFalse</strong>.                                        |
| <strong>msoTrue</strong>           | All characters in the range are formatted as subscript.                                                                         |

Setting the  **SubScript** property to **msoTrue** removes superscript formatting from the text range.


## Example

This example tests the text in the second story and, if it has mixed subscripting, it formats all the text as subscript.


```vb
Sub SubScript() 

 Dim fntSS As Font 

 Set fntSS = Application.ActiveDocument.Stories(2).TextRange.Font 
 With fntSS 
 If .SubScript = msoTriStateMixed Then 
 .SubScript = msoTrue 
 Else 
 MsgBox "Mixed subscript not in this story." 
 End If 
 End With 

End Sub
```


