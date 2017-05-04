---
title: CustomXMLParts Object (Office)
keywords: vbaof11.chm300000
f1_keywords:
- vbaof11.chm300000
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.CustomXMLParts
ms.assetid: 98c1c58e-a08d-6304-8626-1e6705917da3
---


# CustomXMLParts Object (Office)

Represents a collection of  **CustomXMLPart** objects.


## Remarks

There are three default parts that are always created with a document. These are 'Cover pages', 'Doc properties' and 'App properties'. The last two were in previous versions of Microsoft Word but are now provided in XML form in the  **CustomXMLParts** object collection


## Example

The following example adds a node to a  **CustomXMLPart** object that is part of the **CustomXMLParts** object collection.


```
Sub AddPartToCollection() 
    Dim myPart As CustomXMLPart 
 
    Set myPart = ActiveDocument.CustomXMLParts.Add("<author>Mark Twain</author>") 
     
End Sub
```


## Events



|**Name**|
|:-----|
|[PartAfterAdd](http://msdn.microsoft.com/library/c1a263a5-94cb-f563-145b-151a52a31d52%28Office.15%29.aspx)|
|[PartAfterLoad](http://msdn.microsoft.com/library/d59fe837-27b5-300f-133f-ffb01f5f95b9%28Office.15%29.aspx)|
|[PartBeforeDelete](http://msdn.microsoft.com/library/50fa1172-3eac-e091-660e-693a91aaf330%28Office.15%29.aspx)|

## Methods



|**Name**|
|:-----|
|[Add](http://msdn.microsoft.com/library/f2c1588b-c11b-49ca-5db6-4fa4c26d10c5%28Office.15%29.aspx)|
|[SelectByID](http://msdn.microsoft.com/library/e9c0d3a1-c625-bb86-b4ca-6916d4a8a6b0%28Office.15%29.aspx)|
|[SelectByNamespace](http://msdn.microsoft.com/library/39dcce9c-4354-0211-c2cf-393917bf6aef%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/716a8209-ac4f-1cd3-353c-03552ea53035%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/e5c8962f-3f93-8d2c-c5cf-8b485c1b2664%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/b230333f-1bf4-95d6-71d5-089ce884df98%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/801a4462-ccf9-8aa7-f894-4ed89ae09c62%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/6d158523-0297-b823-687c-5b6f3985616b%28Office.15%29.aspx)|

## See also


#### Other resources


[CustomXMLParts Object Members](http://msdn.microsoft.com/library/4e77b5ea-b73c-020f-4abf-25adc200de23%28Office.15%29.aspx)
[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
