
# XMLMapping.CustomXMLNode Property (Word)

 **Last modified:** July 28, 2015

Returns a  **CustomXMLNode** object that represents the custom XML node in the data store to which the content control in the document maps.

## Syntax

 _expression_. **CustomXMLNode**

 _expression_An expression that returns an  ** [XMLMapping](cf76802b-f93d-0f3b-4936-ca357a7d7ff8.md)** object.


## Example

The following example inserts a new content control and custom XML part into the active document, maps the content control to a node in the custom XML part, and then sets the value of the mapped XML node.


```
Dim objCC As ContentControl 
Dim objPart As CustomXMLPart 
Dim objNode As CustomXMLNode 
Dim objMap As XMLMapping 
 
Set objCC = ActiveDocument.ContentControls.Add(wdContentControlText) 
Set objPart = ActiveDocument.CustomXMLParts.Add("<books><book>" &amp; _ 
 "<author></author><title></title><genre></genre><price></price>" &amp; _ 
 "<pub_date></pub_date><abstract></abstract></book></books>") 
 
Set objMap = objCC.XMLMapping 
objMap.SetMapping "/books/book/author", , objPart 
 
Set objNode = objMap.CustomXMLNode 
objNode.Text = "Matt Hink" 
 
objCC.Range.Text = objNode.Text
```


## See also


#### Concepts


 [XMLMapping Object](cf76802b-f93d-0f3b-4936-ca357a7d7ff8.md)
#### Other resources


 [XMLMapping Object Members](8fb27e7a-1d02-4754-87ca-f117cc67cdff.md)
