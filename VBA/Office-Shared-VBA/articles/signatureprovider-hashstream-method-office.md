---
title: SignatureProvider.HashStream Method (Office)
keywords: vbaof11.chm287009
f1_keywords:
- vbaof11.chm287009
ms.prod: office
api_name:
- Office.SignatureProvider.HashStream
ms.assetid: 63f40d22-d49e-d6e8-80d0-7b5c19951b92
ms.date: 06/08/2017
---


# SignatureProvider.HashStream Method (Office)

Allows a signature provider add-in to create a hash value for the document that you can use to determine if the document contents were tampered with after digital signing.


## Syntax

 _expression_. **HashStream**( **_QueryContinue_**, **_Stream_** )

 _expression_ An expression that returns a **SignatureProvider** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _QueryContinue_|Required|**IQueryContinue**|Provides a way to query the host application for permission to continue the hashing process.|
| _Stream_|Required|**IStream**|Contains the data stream.|

### Return Value

Byte


## Remarks

The  **SignatureProvider** object is used exclusively in custom signature provider add-ins. This method is called once per signature data stream in a document. The return value is an array of bytes representing the hash value computed using the hash algorithm.


## Example

The following example gets the hash value of a data stream.


```
 public Array HashStream(object queryContinue, object stream) 
 { 
 using (COMStream comstream = new COMStream(stream)) 
 { 
 using (HashAlgorithm hashalg = HashAlgorithm.Create(this.HashAlgorithmName)) 
 { 
 return hashalg.ComputeHash(comstream); 
 } 
 } 
 } 

```


 **Note**  Signature providers are implemented exclusively in custom COM add-ins and cannot be implemented in Microsoft Visual Basic for Applications (VBA). 


## See also


#### Concepts


[SignatureProvider Object](signatureprovider-object-office.md)
#### Other resources


[SignatureProvider Object Members](signatureprovider-members-office.md)

