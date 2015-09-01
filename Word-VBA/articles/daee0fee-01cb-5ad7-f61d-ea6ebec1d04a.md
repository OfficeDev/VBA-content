
# Range.InsertXML Method (Word)

 **Last modified:** July 28, 2015

Inserts the specified XML into the document at the specified range, replacing any text contained within the range.

## Syntax

 _expression_. **InsertXML**( **_XML_**,  **_Transform_**)

 _expression_An expression that returns a  **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|XML|Required| **String**|Specifies the XML to insert. This can be any valid custom XML.|
|Transform|Optional| **Variant**|Specifies the XML Transformation (XSLT) used to transform the XML. If omitted, the XML is inserted as custom XML without applying a transform.|

### Return Value

Nothing


## Example

The following example inserts the specified XML string into the document at the fifth paragraph. This replaces any text contained within the fifth paragraph.


```
Dim strXML As String 
 
strXML = "<?xml version=""1.0""?><abc:books xmlns:abc=""urn:books"" " &amp; _ 
 "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" " &amp; _ 
 "xsi:schemaLocation=""urn:books books.xsd""><book>" &amp; _ 
 "<author>Matt Hink</author><title>Migration Paths of the Red " &amp; _ 
 "Breasted Robin</title><genre>non-fiction</genre>" &amp; _ 
 "<price>29.95</price><pub_date>2006-05-01</pub_date>" &amp; _ 
 "<abstract>You see them in the spring outside your windows. " &amp; _ 
 "You hear their lovely songs wafting in the warm spring air. " &amp; _ 
 "Now follow their path as they migrate to warmer climes in the fall, " &amp; _ 
 "and then back to your back yard in the spring.</abstract></book></abc:books>" 
 
ActiveDocument.Paragraphs(5).Range.InsertXML strXML
```


## See also


#### Concepts


 [Range Object](15a7a1c4-5f3f-5b6e-60e9-29688de3f274.md)
#### Other resources


 [Range Object Members](3c4a36d9-2a80-5aaf-827b-275a52bfa193.md)
