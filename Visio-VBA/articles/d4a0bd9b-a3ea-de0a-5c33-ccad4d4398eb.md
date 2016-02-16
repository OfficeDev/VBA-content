
# Document.SolutionXMLElementExists Property (Visio)

Indicates whether a named SolutionXML element exists in the document. Read-only.


## Syntax

 _expression_ . **SolutionXMLElementExists**( **_ElementName_** )

 _expression_ A variable that represents a **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ElementName_|Required| **String**|The case-sensitive name of the SolutionXML element.|

### Return Value

Boolean


## Remarks

Because the  **SolutionXMLElement** property can overwrite existing XML data, always use the **SolutionXMLElementExists** property to verify whether _ElementName_ already exists in the document.

