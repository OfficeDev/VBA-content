---
title: Font.BoldBi Property (Publisher)
keywords: vbapb10.chm5373956
f1_keywords:
- vbapb10.chm5373956
ms.prod: publisher
api_name:
- Publisher.Font.BoldBi
ms.assetid: f3a9fa27-6c9c-4d77-0f0d-962afa211d9d
ms.date: 06/08/2017
---


# Font.BoldBi Property (Publisher)

Returns or sets an  **MsoTriState**constant indicating whether the font is bold; used with text in a right-to-left language. Read/write.


## Syntax

 _expression_. **BoldBi**

 _expression_A variable that represents a  **Font** object.


### Return Value

MsoTriState


## Remarks

The  **BoldBi** property value can be one of the following **MsoTriState** constants declared in the Microsoft Office type library.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**|None of the characters in the range are formatted as bold.|
| **msoTriStateMixed**|Return value indicating that the range contains some text formatted as bold and some text not formatted as bold.|
| **msoTriStateToggle**|Set value that switches between  **msoTrue** and **msoFalse**.|
| **msoTrue**|All characters in the range are formatted as bold.|

## Example

This example tests the text in the first story and displays one of two possible messages depending on whether the text is right-to-left formatted and whether its font is bold. For this example to execute properly, there must be at least one story with text in the active publication.


```vb
Sub BoldRtoL() 
 
 Dim stf As Font 
 
 Set stf = Application.ActiveDocument.Stories(1).TextRange.Font 
 
 With stf 
 If .BoldBi = msoTrue Then 
 MsgBox "This story is right-to-left and is bold." 
 Else 
 MsgBox "This story is either not right-to-left" &; _ 
 " or it is not bold." 
 End If 
 End With 
 
End Sub
```


