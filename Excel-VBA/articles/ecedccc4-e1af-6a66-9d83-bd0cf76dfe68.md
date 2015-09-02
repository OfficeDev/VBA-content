
# Pages Object (Excel)

 **Last modified:** July 28, 2015

A collection of pages in a document. Use the  **Pages** collection and the related objects and properties for programmatically defining page layout in a workbook.

## Remarks

Use the  **Pages** property to return a **Pages** collection. The following example accesses all pages in the active worksheet.


```
Dim objPages As Pages 
 
Set objPage = ActiveWorksheet. _ 
 ActiveWindow.Panes(1).Pages
```

Use the  **Item** method to access an individual **Page** object that represents an individual page in a worksheet. The following example accesses the first page in the active worksheet.




```
Dim objPage As Page 
 
Set objPage = ActiveWorksheet.ActiveWindow _ 
 .Panes(1).Pages.Item(1)
```


## See also


#### Concepts


 [Excel Object Model Reference](11ea8598-8a20-92d5-f98b-0da04263bf2c.md)
#### Other resources


 [Pages Object Members](970cda07-ab54-2142-1f0c-d11a1ee4f566.md)
