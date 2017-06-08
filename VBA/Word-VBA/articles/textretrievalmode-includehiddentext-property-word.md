---
title: TextRetrievalMode.IncludeHiddenText Property (Word)
keywords: vbawd10.chm154730498
f1_keywords:
- vbawd10.chm154730498
ms.prod: word
api_name:
- Word.TextRetrievalMode.IncludeHiddenText
ms.assetid: 8904b230-ba07-ecf1-45c3-95d2a11cc434
ms.date: 06/08/2017
---


# TextRetrievalMode.IncludeHiddenText Property (Word)

 **True** if the text retrieved from the specified range includes hidden text. Read/write **Boolean** .


## Syntax

 _expression_ . **IncludeHiddenText**

 _expression_ An expression that returns a **[TextRetrievalMode](textretrievalmode-object-word.md)** object.


## Remarks

The default value is the same as the current setting of the  **Hidden text** option on the **View** tab in the **Options** dialog box ( **Tools** menu) until this property has been set.

 Use the **[Text](find-text-property-word.md)** property with a **[Range](range-object-word.md)** object to retrieve text from the specified range.


## Example

This example displays the text of the first sentence in the active document in a message box. The example uses the  **IncludeHiddenText** property to include hidden text.


```vb
Dim rngTemp As Range 
 
Set rngTemp = ActiveDocument.Sentences(1) 
 
rngTemp.TextRetrievalMode.IncludeHiddenText = True 
MsgBox rngTemp.Text
```

This example posts a message if the entire selection is formatted as hidden text.




```vb
Dim rngTemp As Range 
 
If Selection.Type = wdSelectionNormal Then 
 Set rngTemp = Selection.Range 
 
 rngTemp.TextRetrievalMode.IncludeHiddenText = False 
 If rngTemp.Text = "" Then MsgBox "Selection is hidden" 
End If
```


## See also


#### Concepts


[TextRetrievalMode Object](textretrievalmode-object-word.md)

