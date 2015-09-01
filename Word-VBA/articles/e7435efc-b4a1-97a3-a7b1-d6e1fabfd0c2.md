
# EmailOptions.AutoFormatAsYouTypeApplyTables Property (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


 **True** if Word automatically creates a table when you type a plus sign, a series of hyphens, another plus sign, and so on, and then press ENTER. Read/write **Boolean**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **AutoFormatAsYouTypeApplyTables**

 _expression_A variable that represents an  ** [EmailOptions](41fefa03-c993-e218-0f92-0cf30c0bfbd4.md)** collection.


## Remarks
<a name="sectionSection1"> </a>

The plus signs become the column borders, and the hyphens become the column widths. 


## Example
<a name="sectionSection2"> </a>

This example sets Word to automatically create tables as you type.


```
Options.AutoFormatAsYouTypeApplyTables = True
```

This example returns the status of the  **Tables** option on the **AutoFormat As You Type** tab in the **AutoCorrect** dialog box ( **Tools** menu).




```
Dim blnAutoFormat as Boolean 
 
blnAutoFormat = Options.AutoFormatAsYouTypeApplyTables
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [EmailOptions Object](41fefa03-c993-e218-0f92-0cf30c0bfbd4.md)
#### Other resources


 [EmailOptions Object Members](0f8a549b-283c-dc9d-dc1e-1179a9d6fb0b.md)
