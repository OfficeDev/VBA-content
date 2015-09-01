
# Application.MailMergeAfterMerge Event (Word)

 **Last modified:** July 28, 2015

Occurs after all records in a mail merge have merged successfully.

## Syntax

 _expression_. **Private Sub object_MailMergeAfterMerge**( **_ByVal Doc As Document_**,  **_ByVal DocResult As Document_**)

 _expression_A variable that represents an  ** [Application](d1cf6f8f-4e88-bf01-93b4-90a83f79cb44.md)** object that has been declared with events in a class module. For information about using events with the **Application** object, see [Using Events with the Application Object](784c4c61-7e47-3dbf-46f6-da655f786ca1.md).


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Doc|Required| **Document**|The mail merge main document.|
|DocResult|Required| **Document**|The document created from the mail merge|

## Example

This example displays a message stating that all records in the specified document are finished merging. If the document has been merged to a second document, the message includes the name of the new document. This example assumes that you have declared an application variable called MailMergeApp in your general declarations and have set the variable equal to the Word Application object.


```
Private Sub MailMergeApp_MailMergeAfterMerge(ByVal Doc As Document, _ 
 ByVal DocResult As Document) 
 If DocResult Is Nothing Then 
 MsgBox "Your mail merge on " &amp; _ 
 Doc.Name &amp; " is now finished." 
 
 Else 
 MsgBox "Your mail merge on " &amp; _ 
 Doc.Name &amp; " is now finished and " &amp; _ 
 DocResult.Name &amp; " has been created." 
 End If 
End Sub
```


## See also


#### Concepts


 [Application Object](d1cf6f8f-4e88-bf01-93b4-90a83f79cb44.md)
#### Other resources


 [Application Object Members](71669f1e-65f1-b0f1-b67d-355dfdbebe50.md)
