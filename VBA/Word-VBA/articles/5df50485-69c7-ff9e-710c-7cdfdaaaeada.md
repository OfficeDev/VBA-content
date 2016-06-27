
# TableOfAuthorities.Passim Property (Word)

 **True** if five or more page references to the same authority are replaced with "Passim." Read/write **Boolean** .


## Syntax

 _expression_ . **Passim**

 _expression_ An expression that returns a **[TableOfAuthorities](abd7d600-8b20-0752-4629-8a4f5193dd5d.md)** object.


## Remarks

Corresponds to the \p switch for a Table of Authorities (TOA) field.


## Example

This example formats the first table of authorities in Brief.doc to use page references instead of "Passim."


```vb
Documents("Brief.doc").TablesOfAuthorities(1).Passim = False
```

This example formats the tables of authorities in the active document to replace each instance of five or more page references for the same entry with "Passim."




```vb
For Each myTOA In ActiveDocument.TablesOfAuthorities 
 myToa.Passim = True 
Next myTOA
```


## See also


#### Concepts


[TableOfAuthorities Object](abd7d600-8b20-0752-4629-8a4f5193dd5d.md)
#### Other resources


[TableOfAuthorities Object Members](3e3c6fb0-044b-1b3d-5eff-4be354983675.md)
