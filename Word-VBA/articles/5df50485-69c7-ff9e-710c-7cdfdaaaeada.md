
# TableOfAuthorities.Passim Property (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


 **True** if five or more page references to the same authority are replaced with "Passim." Read/write **Boolean**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **Passim**

 _expression_An expression that returns a  ** [TableOfAuthorities](abd7d600-8b20-0752-4629-8a4f5193dd5d.md)**object.


## Remarks
<a name="sectionSection1"> </a>

Corresponds to the \p switch for a Table of Authorities (TOA) field.


## Example
<a name="sectionSection2"> </a>

This example formats the first table of authorities in Brief.doc to use page references instead of "Passim."


```
Documents("Brief.doc").TablesOfAuthorities(1).Passim = False
```

This example formats the tables of authorities in the active document to replace each instance of five or more page references for the same entry with "Passim."




```
For Each myTOA In ActiveDocument.TablesOfAuthorities 
 myToa.Passim = True 
Next myTOA
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [TableOfAuthorities Object](abd7d600-8b20-0752-4629-8a4f5193dd5d.md)
#### Other resources


 [TableOfAuthorities Object Members](3e3c6fb0-044b-1b3d-5eff-4be354983675.md)
