
# EmailOptions.AutoFormatAsYouTypeReplacePlainTextEmphasis Property (Word)

 **Last modified:** July 28, 2015

 **True** if manual emphasis characters are automatically replaced with character formatting as you type; for example, "*bold*" is changed to " **bold**". Read/write  **Boolean**.

## Syntax

 _expression_. **AutoFormatAsYouTypeReplacePlainTextEmphasis**

 _expression_A variable that represents an  ** [EmailOptions](41fefa03-c993-e218-0f92-0cf30c0bfbd4.md)** collection.


## Example

This example turns on the replacement of manual emphasis characters with character formatting.


```
Options.AutoFormatAsYouTypeReplacePlainTextEmphasis = True
```

This example returns the status of the  ***Bold* and _underline_ with real formatting** option on the **AutoFormat As You Type** tab in the **AutoCorrect** dialog box ( **Tools** menu).




```
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = _ 
 Options.AutoFormatAsYouTypeReplacePlainTextEmphasis
```


## See also


#### Concepts


 [EmailOptions Object](41fefa03-c993-e218-0f92-0cf30c0bfbd4.md)
#### Other resources


 [EmailOptions Object Members](0f8a549b-283c-dc9d-dc1e-1179a9d6fb0b.md)
