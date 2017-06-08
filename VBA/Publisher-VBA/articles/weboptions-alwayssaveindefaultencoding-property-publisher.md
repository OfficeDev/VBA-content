---
title: WebOptions.AlwaysSaveInDefaultEncoding Property (Publisher)
keywords: vbapb10.chm8257539
f1_keywords:
- vbapb10.chm8257539
ms.prod: publisher
api_name:
- Publisher.WebOptions.AlwaysSaveInDefaultEncoding
ms.assetid: e37ff08f-5c09-0a71-27e1-e2a332147087
ms.date: 06/08/2017
---


# WebOptions.AlwaysSaveInDefaultEncoding Property (Publisher)

Returns or sets a  **Boolean** value that specifies whether Web pages within a Web publication should always be saved using default encoding. If **True**, Web pages within a publication will always be saved using the default encoding of the client computer. If  **False**, Web pages will not be saved using default encoding. The default value is  **False**. Read/write.


## Syntax

 _expression_. **AlwaysSaveInDefaultEncoding**

 _expression_A variable that represents a  **WebOptions** object.


### Return Value

Boolean


## Remarks

If the  **AlwaysSaveInDefaultEncoding** property is set to **True** on a given **WebOptions** object, any subsequent attempts to set the **[Encoding](weboptions-encoding-property-publisher.md)** property on that object will be ignored.


## Example

The following example tests whether the Web publication is currently set to be saved using default encoding. If so, the  **AlwaysSaveInDefaultEncoding** property is set to **False**, and the  **Encoding** property is used to set the encoding to Unicode (UTF-8).


```vb
Dim theWO As WebOptions 
 
Set theWO = Application.WebOptions 
 
With theWO 
 If .AlwaysSaveInDefaultEncoding = True Then 
 .AlwaysSaveInDefaultEncoding = False 
 .Encoding = msoEncodingUTF8 
 End If 
End With
```


