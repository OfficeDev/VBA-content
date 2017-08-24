---
title: Font.SuperScript Property (Publisher)
keywords: vbapb10.chm5373972
f1_keywords:
- vbapb10.chm5373972
ms.prod: publisher
api_name:
- Publisher.Font.SuperScript
ms.assetid: 582c02c9-4dcb-f826-8ec0-e9e10702f717
ms.date: 06/08/2017
---


# Font.SuperScript Property (Publisher)

Returns or sets an  **MsoTriState** constant indicating whether characters are formatted as superscript in the specified text range. Read/write.


## Syntax

 _expression_. **SuperScript**

 _expression_A variable that represents a  **Font** object.


### Return Value

MsoTriState


## Remarks

The  **SuperScript** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**| No characters in the range are formatted as superscript.|
| **msoTriStateMixed**|Return value indicating a combination of  **msoTrue** and **msoFalse** for the specified shape range.|
| **msoTriStateToggle**|Set value that switches between  **msoTrue** and **msoFalse**.|
| **msoTrue**|All characters in the range are formatted as superscript.|
Setting the  **SuperScript** property to **msoTrue** removes subscript formatting from the text range.


## Example

This example tests the text in the second story and, if it has mixed superscripting, it formats all the text as superscript.


```vb
Sub SuperScript() 
 
 Dim fntSuper As Font 
 
 Set fntSuper = Application.ActiveDocument.Stories(2).TextRange.Font 
 With fntSuper 
 If .SuperScript = msoTriStateMixed Then 
 .SuperScript = msoTrue 
 Else 
 MsgBox "Mixed superscript not in this story." 
 End If 
 End With 
 
End Sub
```


