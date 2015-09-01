
# Font Object (Word)

 **Last modified:** July 28, 2015

Contains font attributes (such as font name, font size and color) for an object.

## Remarks

Use the  **Font** property to return the **Font** object. The following instruction applies bold formatting to the selection.


```
Selection.Font.Bold = True
```

The following example formats the first paragraph in the active document as 24point Arial and italic.




```
Set myRange = ActiveDocument.Paragraphs(1).Range 
With myRange.Font 
 .Bold = True 
 .Name = "Arial" 
 .Size = 24 
End With
```

The following example changes the formatting of the Heading 2 style in the active document to Arial and bold.




```
With ActiveDocument.Styles(wdStyleHeading2).Font 
 .Name = "Arial" 
 .Italic = True 
End With
```

You can use the  **New** keyword to create a new, stand-alone **Font** object. The following example creates a **Font** object, sets some formatting properties, and then applies the **Font** object to the first paragraph in the active document.




```
Set myFont = New Font 
myFont.Bold = True 
myFont.Name = "Arial" 
ActiveDocument.Paragraphs(1).Range.Font = myFont
```

You can also duplicate a  **Font** object by using the **Duplicate**property. The following example creates a new character style with the character formatting from the selection and italic formatting. The formatting of the selection is not changed.




```
Set aFont = Selection.Font.Duplicate 
aFont.Italic = True 
ActiveDocument.Styles.Add(Name:="Italics", _ 
 Type:=wdStyleTypeCharacter).Font = aFont
```


## See also


#### Concepts


 [Word Object Model Reference](be452561-b436-bb9b-6f94-3faa9a74a6fd.md)
#### Other resources


 [Font Object Members](04a3c706-4062-09bc-70d9-cef3748a7d57.md)
