---
title: WebOptions.EnableIncrementalUpload Property (Publisher)
keywords: vbapb10.chm8257541
f1_keywords:
- vbapb10.chm8257541
ms.prod: publisher
api_name:
- Publisher.WebOptions.EnableIncrementalUpload
ms.assetid: 42d5e22e-7e39-848e-a7e7-5d2019f7e71c
ms.date: 06/08/2017
---


# WebOptions.EnableIncrementalUpload Property (Publisher)

Returns or sets a  **Boolean** value that specifies whether changes made to a Web publication can be uploaded to a Web server independent of the entire publication. If **True**, only changes made to a publication will be uploaded to the Web server when published. If  **False**, the entire publication will be uploaded to the Web server. The default value is  **True**. Read/write.


## Syntax

 _expression_. **EnableIncrementalUpload**

 _expression_A variable that represents an  **WebOptions** object.


### Return Value

Boolean


## Remarks

The  **EnableIncrementalUpload** property applies only to Web publications that have already been published to a Web server. If a Web publication has not already been published to a Web server, the entire publication will be published to the server during the initial publishing process, regardless of whether the **EnableIncrementalUpload** property is set to **True**. If a Web publication has already been published to a Web server and the  **EnableIncrementalUpload** property is then set to **True**, only changes made to the Web publication, and not the entire publication, after this point will be published to the server.


## Example

The following example tests whether the Web publication is set to upload only changes made to the publication. If not, the  **EnableIncrementalUpload** property is set to **True** to specify that only changes to the publication be uploaded to the Web server.


```vb
Dim theWO As WebOptions 
 
Set theWO = Application.WebOptions 
 
With theWO 
 If .EnableIncrementalUpload = False Then 
 .EnableIncrementalUpload = True 
 End If 
End With
```


