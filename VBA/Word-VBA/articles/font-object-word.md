---
title: Font Object (Word)
keywords: vbawd10.chm2386
f1_keywords:
- vbawd10.chm2386
ms.prod: word
api_name:
- Word.Font
ms.assetid: bc97f4df-fc81-d6c8-e99a-d50dc793b7ae
ms.date: 06/08/2017
---


# Font Object (Word)

Contains font attributes (such as font name, font size and color) for an object.


## Remarks

Use the  **Font** property to return the **Font** object. The following instruction applies bold formatting to the selection.


```vb
Selection.Font.Bold = True
```

The following example formats the first paragraph in the active document as 24point Arial and italic.




```vb
Set myRange = ActiveDocument.Paragraphs(1).Range 
With myRange.Font 
 .Bold = True 
 .Name = "Arial" 
 .Size = 24 
End With
```

The following example changes the formatting of the Heading 2 style in the active document to Arial and bold.




```vb
With ActiveDocument.Styles(wdStyleHeading2).Font 
 .Name = "Arial" 
 .Italic = True 
End With
```

You can use the  **New** keyword to create a new, stand-alone **Font** object. The following example creates a **Font** object, sets some formatting properties, and then applies the **Font** object to the first paragraph in the active document.




```vb
Set myFont = New Font 
myFont.Bold = True 
myFont.Name = "Arial" 
ActiveDocument.Paragraphs(1).Range.Font = myFont
```

You can also duplicate a  **Font** object by using the **Duplicate** property. The following example creates a new character style with the character formatting from the selection and italic formatting. The formatting of the selection is not changed.




```vb
Set aFont = Selection.Font.Duplicate 
aFont.Italic = True 
ActiveDocument.Styles.Add(Name:="Italics", _ 
 Type:=wdStyleTypeCharacter).Font = aFont
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

