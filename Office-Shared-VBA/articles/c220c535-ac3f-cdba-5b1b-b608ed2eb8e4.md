
# CustomXMLPart.SelectNodes Method (Office)

 **Last modified:** July 28, 2015

Selects a collection of nodes from a custom XML part.

## Syntax

 _expression_. **SelectNodes**( **_XPath_**)

 _expression_An expression that returns a  **CustomXMLPart** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|XPath|Required| **String**|Contains the XPath expression.|

### Return Value

CustomXMLNodes


## Example

The following example demonstrates adding a custom XML part, selecting a part matching a namespace URI, and then selecting nodes within that part that match an XPath expression.


```
Dim cxp1 As CustomXMLPart 
Dim cxn As CustomXMLNode 
 
' Add a custom xml part. 
ActiveDocument.CustomXMLParts.Add "<supplier>" 
 
' Return the first custom xml part with the given namespace. 
Set cxp1 = ActiveDocument.CustomXMLParts("urn:invoice:namespace")  
 
' Get all of the nodes matching an XPath expression. 
 Set cxns = cxp1.SelectNodes("//*[@unitPrice > 20]") 

```


## See also


#### Concepts


 [CustomXMLPart Object](a4f90bac-01d6-bba4-f64b-a64e2b122cfd.md)
#### Other resources


 [CustomXMLPart Object Members](76fe85f4-5a35-7d12-2989-6f17a094dcdf.md)
