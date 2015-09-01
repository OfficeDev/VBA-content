
# Options.TabIndentKey Property (Word)

 **Last modified:** July 28, 2015

 **True** if the TAB and BACKSPACE keys can be used to increase and decrease, respectively, the left indent of paragraphs and if the BACKSPACE key can be used to change right-aligned paragraphs to centered paragraphs and centered paragraphs to left-aligned paragraphs. Read/write **Boolean**.

## Syntax

 _expression_. **TabIndentKey**

 _expression_Required. A variable that represents an  ** [Options](873b7b99-3fe1-fd89-9ece-a9355cb827dc.md)** collection.


## Example

This example sets Word so that the TAB and BACKSPACE keys set the left indent instead of inserting and deleting tabs.


```
Options.TabIndentKey = True
```


## See also


#### Concepts


 [Options Object](873b7b99-3fe1-fd89-9ece-a9355cb827dc.md)
#### Other resources


 [Options Object Members](76cd9dfe-6bbb-4c3d-0bfc-79a62bedd15e.md)
