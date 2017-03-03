---
title: MetaProperties Object (Office)
keywords: vbaof11.chm274000
f1_keywords:
- vbaof11.chm274000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.MetaProperties
ms.assetid: 957a6e06-3348-b180-3655-06ffbfb69e12
---


# MetaProperties Object (Office)

Represents a collection of properties describing the metadata stored in a document.


## Example

In the following example, a  **MetaProperties** object is passed to a validation function. The function then validates the value of a single property represented by its index and returns the result.


```
Function ValidateMetaProperty(ByVal metaProps As MetaProperties) As String 
Dim result As String 
 
result = metaProps(1).Validate 
ValidateMetaProperty = result 
End Function
```


## Methods



|**Name**|
|:-----|
|[GetItemByInternalName](http://msdn.microsoft.com/library/27c6bcd8-8631-1dbe-5df1-67c33b757c03%28Office.15%29.aspx)|
|[Validate](http://msdn.microsoft.com/library/658532c6-c8c0-ff01-3736-4161a09af2bb%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/40f520da-9408-06f9-f51d-1b4dda0d452b%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/ceb7c117-4d5a-511c-a849-b3cc9041d298%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/377c8cee-9561-21aa-666c-f5e291ca899a%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/e1c30443-08c3-85bc-bfdd-59cd825b63e5%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/cafd45a4-59ea-4459-3c35-75062964e5c9%28Office.15%29.aspx)|
|[SchemaXml](http://msdn.microsoft.com/library/c51acc59-3014-8678-c697-425be9dc3aeb%28Office.15%29.aspx)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
[MetaProperties Object Members](http://msdn.microsoft.com/library/0e2efa13-130c-59ad-07ee-8499f502064a%28Office.15%29.aspx)
