
# BuildingBlocks Object (Word)

 **Last modified:** July 28, 2015

Represents a collection of  ** [BuildingBlock](2558b89f-8552-bb71-fa40-101cab2635ba.md)** objects for a specific building block type and category in a template.

## Remarks

Use the  ** [Add](22725f33-4de0-95cd-d4a5-a2379b0130c4.md)** method to create a new building block and add it to a template. The following example adds the selected text to the watermarks building block gallery of the first template in the ** [Templates](de62f768-011a-7446-48c3-1c4512da5f7c.md)** collection.


```
Dim objTemplate As Template 
Dim objBB As BuildingBlock 
 
Set objTemplate = Templates(1) 
 
Set objBB = objTemplate.BuildingBlockEntries _ 
 .Add(Name:="New Building Block Entry", _ 
 Type:=wdTypeWatermarks, _ 
 Category:="General", _ 
 Range:=Selection.Range)
```

The collection returned with the  **BuildingBlocks** collection is a filtered collection based on the type and category. Depending on how you access the collection, the collection returned changes. For example, if you access a collection of building blocks with a type of **wdTypeAutoText** with a category of "General", the returned collection may be different from the collection returned if you access a collection of building blocks with a type of **wdTypeAutoText** with a category of "Custom". It is also different from the collection returned if you access the collection of building blocks with a type of **wdTypeCustomAutoText** with a category of "General".

For more information about building blocks, see  [Working with Building Blocks](c32a8972-a6fc-bb66-b62a-039b88580b37.md).


## See also


#### Concepts


 [Word Object Model Reference](be452561-b436-bb9b-6f94-3faa9a74a6fd.md)
#### Other resources


 [BuildingBlocks Object Members](865639de-1856-d542-fe6b-a09425c050f0.md)
