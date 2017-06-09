---
title: Font.SmallCaps Property (Publisher)
keywords: vbapb10.chm5373971
f1_keywords:
- vbapb10.chm5373971
ms.prod: publisher
api_name:
- Publisher.Font.SmallCaps
ms.assetid: ab50b850-f371-7d8e-0c19-00ad68e700f0
ms.date: 06/08/2017
---


# Font.SmallCaps Property (Publisher)

Returns or sets an  **MsoTriState** constant indicating whether the specified text is formatted as small caps. Read/write.


## Syntax

 _expression_. **SmallCaps**

 _expression_A variable that represents a  **Font** object.


### Return Value

MsoTriState


## Remarks

The  **SmallCaps** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**|None of the characters in the range are formatted as small caps.|
| **msoTriStateMixed**|Return value indicating a combination of  **msoTrue** and **msoFalse** for the specified shape range.|
| **msoTriStateToggle**|Set value that switches between  **msoTrue** and **msoFalse**.|
| **msoTrue**| All of the characters in the range are formatted as small caps.|
Setting the  **SmallCaps** property to **msoTrue** removes all caps formatting from the text range.


## Example

This example tests the text in the second story and, if it has mixed small caps formatting, it formats all the text as small caps.


```vb
Sub SmallCap() 
 
 Dim fntSC As Font 
 
 Set fntSC = Application.ActiveDocument.Stories(2).TextRange.Font 
 With fntSC 
 If .SmallCaps = msoTriStateMixed Then 
 .SmallCaps = msoTrue 
 Else 
 MsgBox "Mixed small caps are not in this story." 
 End If 
 End With 
 
End Sub
```


