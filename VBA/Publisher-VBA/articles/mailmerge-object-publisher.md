---
title: MailMerge Object (Publisher)
keywords: vbapb10.chm6291455
f1_keywords:
- vbapb10.chm6291455
ms.prod: publisher
api_name:
- Publisher.MailMerge
ms.assetid: 028e1e42-c61c-9b2b-4aec-d6a184504ec1
ms.date: 06/08/2017
---


# MailMerge Object (Publisher)

Represents the mail merge and catalog merge functionality in Microsoft Publisher.


## Example

Use the  **[MailMerge](http://msdn.microsoft.com/library/15b1a8aa-3472-c67d-1d99-92617b05c157%28Office.15%29.aspx)** property to return the **MailMerge** object. The **MailMerge** object is always available regardless of whether the mail merge or catalog merge operation has begun. The following example merges and prints the main publication with the first three records in the attached data source.


```
Sub SelectiveMerge() 
 Dim mrgMain As MailMerge 
 Set mrgMain = ActiveDocument.MailMerge 
 With mrgMain.DataSource 
 .FirstRecord = 1 
 .LastRecord = 3 
 End With 
 mrgMain.Execute True 
End Sub
```


## Methods



|**Name**|
|:-----|
|[CreateShortcut](http://msdn.microsoft.com/library/96878925-41ce-4873-931e-d5c05307a94a%28Office.15%29.aspx)|
|[Execute](http://msdn.microsoft.com/library/edcabcc5-f2ce-53ce-d422-0d6fcb5f8a33%28Office.15%29.aspx)|
|[ExportRecipientList](http://msdn.microsoft.com/library/230d0f66-7368-51b7-8233-3fd54cfd0fe4%28Office.15%29.aspx)|
|[OpenDataSource](http://msdn.microsoft.com/library/4473e566-687f-595e-9fd6-a5483021cb48%28Office.15%29.aspx)|
|[ShowWizardEx](http://msdn.microsoft.com/library/3815204f-5f09-5a25-a2e4-5de4889c9919%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/44a89300-ff8a-ccc6-5646-6ef7e4cb8138%28Office.15%29.aspx)|
|[DataSource](http://msdn.microsoft.com/library/19b32513-fd57-617a-38e2-6230e3e036b9%28Office.15%29.aspx)|
|[DocumentUpdating](http://msdn.microsoft.com/library/c65ca4a0-e5eb-d97e-9126-4af86f4e805f%28Office.15%29.aspx)|
|[EmailMergeEnvelope](http://msdn.microsoft.com/library/96ddcd72-c87f-9ddb-5a7f-b91be715fc79%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/983636d1-f748-1f47-a52d-8c44c820de16%28Office.15%29.aspx)|
|[SuppressBlankLines](http://msdn.microsoft.com/library/3b41e0c0-8588-e86a-77ed-90c4692c03dc%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/cd31c23f-4059-c6ae-851a-ec9b7f107724%28Office.15%29.aspx)|
|[ViewMailMergeFieldCodes](http://msdn.microsoft.com/library/05b5e6e2-10ae-c6e0-3214-7016295703e2%28Office.15%29.aspx)|
|[WizardState](http://msdn.microsoft.com/library/a237cb3f-2c03-5f62-fa67-d4aa7703389d%28Office.15%29.aspx)|

