---
title: CustomXMLValidationError Object (Office)
keywords: vbaof11.chm307000
f1_keywords:
- vbaof11.chm307000
ms.prod: office
api_name:
- Office.CustomXMLValidationError
ms.assetid: 7f7ced9a-0878-9287-fe66-a7f0ffdc45b6
ms.date: 06/08/2017
---


# CustomXMLValidationError Object (Office)

Represents a single validation error in a  **CustomXMLValidationErrors** collection.


## Remarks

Validation errors can either be triggered when validating an operation against the schema, such as when adding a node, or triggered by the user when some condition is not met. For example, if a start date is later than an end date. 


## Example

The following example adds a custom part and then adds a child node to that part. Any errors that occur are added to the  **CustomXMLValidationErrors** collection and then displayed in the Debug window.


```
Dim ValErrors As CustomXMLValidationErrors 
Dim ValError As CustomXMLValidationError 
Dim cxp1 As CustomXMLPart 
Dim intError As Integer 
 
On Error Go To validation_error 
 
 With ActiveDocument 
 
    ' Add and populate a custom xml part 
    set cxp1 = .CustomXMLParts.Add "<invoice>" 
 
    ' Add a node 
    cxp1.AddNode "<quantity>", "supplier", "urn:invoice:namespace" 
 
 End With 
 
If ValErrors.Count > 0 then 
   For Each ValError In ValErrors 
      DeBug.Print("Error name: " &amp; ValError.Name &amp; " Error description: " &amp; ValError.Text)  
   Next 
End If 
 
Exit Sub 
 
validation_error: 
   CustomXMLValidationErrors.Add(ValError.Name, ValError.Text)) 
Resume
```


## Methods



|**Name**|
|:-----|
|[Delete](customxmlvalidationerror-delete-method-office.md)|

## Properties



|**Name**|
|:-----|
|[Application](customxmlvalidationerror-application-property-office.md)|
|[Creator](customxmlvalidationerror-creator-property-office.md)|
|[ErrorCode](customxmlvalidationerror-errorcode-property-office.md)|
|[Name](customxmlvalidationerror-name-property-office.md)|
|[Node](customxmlvalidationerror-node-property-office.md)|
|[Parent](customxmlvalidationerror-parent-property-office.md)|
|[Text](customxmlvalidationerror-text-property-office.md)|
|[Type](customxmlvalidationerror-type-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
