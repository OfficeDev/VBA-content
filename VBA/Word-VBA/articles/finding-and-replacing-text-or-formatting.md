---
title: Finding and Replacing Text or Formatting
ms.prod: word
ms.assetid: 9ab9f4a7-9833-5a78-56b0-56a161480f18
ms.date: 06/08/2017
---


# Finding and Replacing Text or Formatting

Finding and replacing is exposed by the  **[Find](find-object-word.md)** and  **[Replacement](replacement-object-word.md)** objects. The  **Find** object is available from the **[Selection](selection-object-word.md)** object and the  **[Range](range-object-word.md)** object. The find action differs slightly depending upon whether you access the  **Find** object from the **Selection** object or the **Range** object.


## Finding text and selecting it

If the  **Find** object is accessed from the **Selection** object, the selection is changed when the find criteria is found. The following example selects the next occurrence of the word "Hello." If the end of the document is reached before the word "Hello" is found, the search is stopped.


```vb
With Selection.Find 
 .Forward = True 
 .Wrap = wdFindStop 
 .Text = "Hello" 
 .Execute 
End With
```

The  **Find** object includes properties that relate to the options in the **Find and Replace** dialog box. You can set the individual properties of the **Find** object or use arguments with the **[Execute](find-execute-method-word.md)** method, as shown in the following example.




```
Selection.Find.Execute FindText:="Hello", _ 
 Forward:=True, Wrap:=wdFindStop
```


## Finding text without changing the selection

If the  **Find** object is accessed from a **Range** object, the selection is not changed but the **Range** is redefined when the find criteria is found. The following example locates the first occurrence of the word "blue" in the active document. If the find operation is successful, the range is redefined and bold formatting is applied to the word "blue."


```vb
With ActiveDocument.Content.Find 
 .Text = "blue" 
 .Forward = True 
 .Execute 
 If .Found = True Then .Parent.Bold = True 
End With
```

The following example performs the same result as the previous example, using arguments of the  **Execute** method.




```vb
Set myRange = ActiveDocument.Content 
myRange.Find.Execute FindText:="blue", Forward:=True 
If myRange.Find.Found = True Then myRange.Bold = True
```


## Using the Replacement object

The  **Replacement** object represents the replace criteria for a find and replace operation. The properties and methods of the **Replacement** object correspond to the options in the **Find and Replace** dialog box ( **Edit** menu).

The  **Replacement** object is available from the **Find** object. The following example replaces all occurrences of the word "hi" with "hello". The selection changes when the find criteria is found because the **Find** object is accessed from the **Selection** object.




```vb
With Selection.Find 
 .ClearFormatting 
 .Text = "hi" 
 .Replacement.ClearFormatting 
 .Replacement.Text = "hello" 
 .Execute Replace:=wdReplaceAll, Forward:=True, _ 
 Wrap:=wdFindContinue 
End With
```

The following example removes bold formatting in the active document. The  **[Bold](font-bold-property-word.md)** property is  **True** for the **Find** object and **False** for the **Replacement** object. To find and replace formatting, set the find and replace text to empty strings ("") and set the **_Format_** argument of the **Execute** method to **True**. The selection remains unchanged because the  **Find** object is accessed from a **Range** object (the **[Content](document-content-property-word.md)** property returns a  **Range** object).




```vb
With ActiveDocument.Content.Find 
 .ClearFormatting 
 .Font.Bold = True 
 With .Replacement 
 .ClearFormatting 
 .Font.Bold = False 
 End With 
 .Execute FindText:="", ReplaceWith:="", _ 
 Format:=True, Replace:=wdReplaceAll 
End With
```


