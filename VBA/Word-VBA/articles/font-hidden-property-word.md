---
title: Font.Hidden Property (Word)
keywords: vbawd10.chm156369028
f1_keywords:
- vbawd10.chm156369028
ms.prod: word
api_name:
- Word.Font.Hidden
ms.assetid: a857f5e5-cda6-9402-dc82-6ed3bd93e2c4
ms.date: 06/08/2017
---


# Font.Hidden Property (Word)

 **True** if the font is formatted as hidden text. Read/write **Long** .


## Syntax

 _expression_ . **Hidden**

 _expression_ Required. A variable that represents a **[Font](font-object-word.md)** object.


## Remarks

This property returns  **True** , **False** or **wdUndefined** (a mixture of **True** and **False** ). Can be set to **True** , **False** , or **wdToggle** .

To control the display of hidden text, use the  **ShowHiddenText** property of the **View** object.

To control whether properties and methods that return  **Range** objects include or exclude hidden text when hidden text isn't displayed, use the **IncludeHiddenText** property of the **TextRetrievalMode** object.


## Example

This example checks the selection for hidden text.


```vb
If Selection.Type = wdSelectionNormal Then 
 If Selection.Font.Hidden = wdUndefined or _ 
 Selection.Font.Hidden = True Then 
 MsgBox "There is hidden text in the selection." 
 Else 
 MsgBox "No hidden text in the selection." 
 End If 
Else 
 MsgBox "You need to select some text." 
End If
```

This example makes all hidden text in the active window visible and then formats the selection as hidden text.




```vb
ActiveDocument.ActiveWindow.View.ShowHiddenText = True 
If Selection.Type = wdSelectionNormal Then _ 
 Selection.Font.Hidden = True
```


## See also


#### Concepts


[Font Object](font-object-word.md)

