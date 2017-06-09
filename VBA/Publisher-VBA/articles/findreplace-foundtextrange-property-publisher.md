---
title: FindReplace.FoundTextRange Property (Publisher)
keywords: vbapb10.chm8323075
f1_keywords:
- vbapb10.chm8323075
ms.prod: publisher
api_name:
- Publisher.FindReplace.FoundTextRange
ms.assetid: 8d0d3177-2d32-7df6-8b88-b354ec0a3d7b
ms.date: 06/08/2017
---


# FindReplace.FoundTextRange Property (Publisher)

Returns a  **TextRange** object that represents the found text or replaced text of a find operation. Read-only.


## Syntax

 _expression_. **FoundTextRange**

 _expression_A variable that represents a  **FindReplace** object.


### Return Value

TextRange


## Remarks

The actual  **TextRange** object returned by the **FoundTextRange** property is determined by the value of the **ReplaceScope** property. The following table lists the corresponding values of these properties.



|for  **ReplaceScope** = **pbReplaceScopeAll**| **FoundTextRange** = Empty|
|for  **ReplaceScope** = **pbReplaceScopeNone**| **FoundTextRange** = Find text range|
|for  **ReplaceScope** = **pbReplaceScopeOne**| **FoundTextRange** = Replace text range|
When  **ReplaceScope** is set to **pbReplaceScopeAll**, the  **FoundTextRange** property is empty. Any attempt to access it returns "Access Denied." The way to manipulate the text range of the searched text is to set the **ReplaceScope** property to **pbReplaceScopeNone** or **pbReplaceScopeOne** and access the text range of the searched or replaced text for each occurrence found.


## Example

When  **ReplaceScope** is set to **pbReplaceScopeNone**,  **FoundTextRange** returns the text range of the searched text. The following example illustrates how the font attributes of the find text range can be accessed when **ReplaceScope** is set to **pbReplaceScopeNone**.


```vb
With TextRange.Find 
 .Clear 
 .FindText = "important" 
 .ReplaceScope = pbReplaceScopeNone 
 Do While .Execute = True 
 'The FoundTextRange contains the word "important". 
 If .FoundTextRange.Font.Italic = msoFalse Then 
 .FoundTextRange.Font.Italic = msoTrue 
 End If 
 Loop 
End With
```

When  **ReplaceScope** is set to **pbReplaceScopeOne**, the text range of the searched text is replaced. Therefore, the  **FoundTextRange** property returns the text range of the replacement text. The following example demonstrates how the font attributes of the replaced text range can be accessed when **ReplaceScope** is set to **pbReplaceScopeOne**. 




```vb
With Document.Find 
 .Clear 
 .FindText = "important" 
 .ReplaceWithText = "urgent" 
 .ReplaceScope = pbReplaceScopeOne 
 Do While .Execute = True 
 'The FoundTextRange contains the word "urgent". 
 If .FoundTextRange.Font.Bold = msoFalse Then 
 .FoundTextRange.Font.Bold = msoTrue 
 End If 
 Loop 
End With
```

This example replaces each example of the word "bizarre" with the word "strange" and applies italic formatting and bold formatting to the replaced text. 




```vb
Dim objDocument As Document 
 
Set objDocument = ActiveDocument 
With objDocument.Find 
 .Clear 
 .FindText = "bizarre" 
 .ReplaceWithText = "strange" 
 .ReplaceScope = pbReplaceScopeOne 
 Do While .Execute = True 
 .FoundTextRange.Font.Italic = msoTrue 
 .FoundTextRange.Font.Bold = msoTrue 
 Loop 
End With
```

This example finds all occurrences of the word "important" and applies italic formatting to it.




```vb
Dim objTextRange As TextRange 
 
Set objTextRange = ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange 
With objTextRange.Find 
 .Clear 
 .FindText = "important" 
 .ReplaceScope = pbReplaceScopeNone 
 Do While .Execute = True 
 .FoundTextRange.Font.Italic = msoTrue 
 Loop 
End With
```


