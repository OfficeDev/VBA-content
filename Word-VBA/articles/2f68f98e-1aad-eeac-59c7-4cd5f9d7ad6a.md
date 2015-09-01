
# Application.NewDocument Property (Word)

 **Last modified:** July 28, 2015

Returns a  **NewFile** object that represents a document listed on the **New** tab.

## Syntax

 _expression_. **NewDocument**

 _expression_A variable that represents an  ** [Application](d1cf6f8f-4e88-bf01-93b4-90a83f79cb44.md)** object.


## Example

This example creates a document list item on the New Document task pane in the New From Existing File section.


```
Sub CreateNewDocument() 
 Application.NewDocument.Add FileName:="C:\NewFile.doc", _ 
 Section:=msoNewfromExistingFile, DisplayName:="New File", _ 
 Action:=msoCreateNewFile 
End Sub
```


## See also


#### Concepts


 [Application Object](d1cf6f8f-4e88-bf01-93b4-90a83f79cb44.md)
#### Other resources


 [Application Object Members](71669f1e-65f1-b0f1-b67d-355dfdbebe50.md)
