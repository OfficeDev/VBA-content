
# Index.NumberOfColumns Property (Word)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Sets or returns the number of columns for each page of an index. Read/write  **Long**.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **NumberOfColumns**

 _expression_An expression that an  ** [Index](6a2aab98-485b-01c3-8d9b-9e108b455e22.md)**object.


## Remarks
<a name="sectionSection1"> </a>

Specifying 0 (zero) sets the number of columns in the index to the same number as in the document.


## Example
<a name="sectionSection2"> </a>

This example sets the number of columns in the first index to the same number as in the active document.


```
ActiveDocument.Indexes(1).NumberOfColumns = 0
```

This example sets a two-column format for each index in the active document.




```
For Each myIndex In ActiveDocument.Indexes 
 myIndex.NumberOfColumns = 2 
Next myIndex
```


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [Index Object](6a2aab98-485b-01c3-8d9b-9e108b455e22.md)
#### Other resources


 [Index Object Members](de9f0a3c-dd30-84bd-e122-2d20fa6b3d37.md)
