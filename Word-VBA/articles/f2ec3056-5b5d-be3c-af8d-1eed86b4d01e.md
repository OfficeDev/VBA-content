
# Font.Outline Property (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


 **True** if the font is formatted as outline. Read/write **Long**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Outline**

 _expression_An expression that returns a  ** [Font](bc97f4df-fc81-d6c8-e99a-d50dc793b7ae.md)**object.


## Remarks
<a name="sectionSection1"> </a>

Returns  **True**,  **False**, or  **wdUndefined** (a mixture of **True** and **False**). Can be set to  **True**,  **False**, or  **wdToggle**.


## Example
<a name="sectionSection2"> </a>

This example applies outline font formatting to the first three words in the active document.


```
Set myRange = ActiveDocument.Range(Start:= _ 
 ActiveDocument.Words(1).Start, _ 
 End:=ActiveDocument.Words(3).End) 
myRange.Font.Outline = True
```

This example toggles outline formatting for the selected text.




```
Selection.Font.Outline = wdToggle
```

This example removes outline font formatting from the selection if outline formatting is partially applied to the selection.




```
Set myFont = Selection.Font 
If myFont.Outline = wdUndefined Then 
 myFont.Outline = False 
End If
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Font Object](bc97f4df-fc81-d6c8-e99a-d50dc793b7ae.md)
#### Other resources


 [Font Object Members](04a3c706-4062-09bc-70d9-cef3748a7d57.md)
