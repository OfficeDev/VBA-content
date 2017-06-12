---
title: TextRetrievalMode.IncludeFieldCodes Property (Word)
keywords: vbawd10.chm154730499
f1_keywords:
- vbawd10.chm154730499
ms.prod: word
api_name:
- Word.TextRetrievalMode.IncludeFieldCodes
ms.assetid: 9055d78b-ddf4-3e58-a42d-813ef838cdf2
ms.date: 06/08/2017
---


# TextRetrievalMode.IncludeFieldCodes Property (Word)

 **True** if the text retrieved from the specified range includes field codes. Read/write **Boolean** .


## Syntax

 _expression_ . **IncludeFieldCodes**

 _expression_ An expression that returns a **[TextRetrievalMode](textretrievalmode-object-word.md)** object.


## Remarks

The default value is the same as the setting of the  **Field codes** option on the **View** tab in the **Options** dialog box ( **Tools** menu) until this property has been set.

Use the  **[Text](find-text-property-word.md)** property with a **[Range](range-object-word.md)** object to retrieve text from the specified range.


## Example

This example displays the text of the first paragraph in the active document in a message box. The example uses the  **IncludeFieldCodes** property to exclude field codes.


```vb
Dim rngTemp As Range 
 
Set rngTemp = ActiveDocument.Paragraphs(1).Range 
 
rngTemp.TextRetrievalMode.IncludeFieldCodes = False 
MsgBox rngTemp.Text
```

This example excludes field codes and hidden text from the range that refers to the selected text, and then it displays the text in a message box.




```vb
Dim rngTemp As Range 
 
If Selection.Type = wdSelectionNormal Then 
 Set rngTemp = Selection.Range 
 With rngTemp.TextRetrievalMode 
 .IncludeHiddenText = False 
 .IncludeFieldCodes = False 
 End With 
 MsgBox rngTemp.Text 
End If
```


## See also


#### Concepts


[TextRetrievalMode Object](textretrievalmode-object-word.md)

