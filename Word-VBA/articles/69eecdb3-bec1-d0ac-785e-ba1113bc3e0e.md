
# Options.AllowAccentedUppercase Property (Word)

 **True** if accents are retained when a French language character is changed to uppercase. Read/write **Boolean** .


## Syntax

 _expression_ . **AllowAccentedUppercase**

 _expression_ A variable that represents a **[Options](873b7b99-3fe1-fd89-9ece-a9355cb827dc.md)** object.


## Remarks

This property affects only text that's been marked as standard French. For all other languages, accents are always retained even if the  **AllowAccentedUppercase** property is set to **False** .

If you change a character back to lowercase after an accent mark has been stripped from it, the accent won't reappear.


## Example

This example sets Word to remove accent marks when characters in French text are changed to uppercase.


```vb
Options.AllowAccentedUppercase = False
```

This example returns the status of the Allow accented uppercase in French option on the Edit tab in the Options dialog box.




```vb
Dim blnUppercaseAccents as Boolean 
 
blnUppercaseAccents = Options.AllowAccentedUppercase
```


## See also


#### Concepts


[Options Object](873b7b99-3fe1-fd89-9ece-a9355cb827dc.md)
#### Other resources


[Options Object Members](76cd9dfe-6bbb-4c3d-0bfc-79a62bedd15e.md)
