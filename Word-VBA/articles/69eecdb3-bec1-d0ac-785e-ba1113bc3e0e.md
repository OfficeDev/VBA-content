
# Options.AllowAccentedUppercase Property (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


 **True** if accents are retained when a French language character is changed to uppercase. Read/write **Boolean**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **AllowAccentedUppercase**

 _expression_A variable that represents a  ** [Options](873b7b99-3fe1-fd89-9ece-a9355cb827dc.md)** object.


## Remarks
<a name="sectionSection1"> </a>

This property affects only text that's been marked as standard French. For all other languages, accents are always retained even if the  **AllowAccentedUppercase** property is set to **False**.

If you change a character back to lowercase after an accent mark has been stripped from it, the accent won't reappear.


## Example
<a name="sectionSection2"> </a>

This example sets Word to remove accent marks when characters in French text are changed to uppercase.


```
Options.AllowAccentedUppercase = False
```

This example returns the status of the Allow accented uppercase in French option on the Edit tab in the Options dialog box.




```
Dim blnUppercaseAccents as Boolean 
 
blnUppercaseAccents = Options.AllowAccentedUppercase
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Options Object](873b7b99-3fe1-fd89-9ece-a9355cb827dc.md)
#### Other resources


 [Options Object Members](76cd9dfe-6bbb-4c3d-0bfc-79a62bedd15e.md)
