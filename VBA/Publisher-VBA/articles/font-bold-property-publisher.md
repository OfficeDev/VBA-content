---
title: Font.Bold Property (Publisher)
keywords: vbapb10.chm5373955
f1_keywords:
- vbapb10.chm5373955
ms.prod: publisher
api_name:
- Publisher.Font.Bold
ms.assetid: 3b9ba2b0-c319-9d08-9a36-5b292046962e
ms.date: 06/08/2017
---


# Font.Bold Property (Publisher)

Returns or sets an  **MsoTriState**constant that represents the state of the  **Bold** property on the characters in a text range. Read/write.


## Syntax

 _expression_. **Bold**

 _expression_A variable that represents a  **Font** object.


### Return Value

MsoTriState


## Remarks

The  **Bold** property value can be one of the following **MsoTriState** constants declared in the Microsoft Office type library.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**|None of the characters in the range are formatted as bold.|
| **msoTriStateMixed**|Return value indicating that the range contains some text formatted as bold and some text not formatted as bold.|
| **msoTriStateToggle**|Set value that switches between  **msoTrue** and **msoFalse**.|
| **msoTrue**|All characters in the range are formatted as bold.|

## Example

This example tests all the text in the second story of the active publication and, if it contains both bold text and not bold text, it sets all the text to bold. If the text is all bold or all not bold, a message is displayed informing the user there is no mixed bolding. For this code to execute properly, there must be two or more stories with text in the active publication.


```vb
Sub BoldStory() 
 
 Dim stf As Publisher.Font 
 
 Set stf = Application.ActiveDocument.Stories(2).TextRange.Font 
 With stf 
 If .Bold = msoTriStateMixed Then 
 .Bold = msoTrue 
 Else 
 MsgBox "Mixed bolding is not in this story." 
 End If 
 End With 
 
End Sub
```


