---
title: "Объект MailMergeMappedDataField (издатель)"
keywords: vbapb10.chm6619135
f1_keywords: vbapb10.chm6619135
ms.prod: publisher
api_name: Publisher.MailMergeMappedDataField
ms.assetid: 3711d28e-f005-27fb-88b5-8674d4ece887
ms.date: 06/08/2017
ms.openlocfilehash: 6db8a4f2b18e12efa4a8fb3b8c69942257bb31f1
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergemappeddatafield-object-publisher"></a>Объект MailMergeMappedDataField (издатель)

Представляет одного сопоставленные поля данных. Объект **MailMergeMappedDataField** является элементом коллекции **[MailMergeMappedDataFields](mailmergemappeddatafields-object-publisher.md)** . Сопоставленные данные поля — поля, содержащиеся в Microsoft Publisher, представляющий часто используемых имя или адрес информации, например, имя. Если источник данных содержит имя поля или вариантов (например, имя, FirstName, во-первых или отображающее общие сведения), в соответствующее поле сопоставленные данные автоматически сопоставляются поля в источнике данных. При публикации для объединения с более одного источника данных, сопоставленные поля данных устраняют на повторный ввод поля в публикации для подтверждения с именами полей в базе данных.
 


## <a name="example"></a>Пример

Используйте **MappedDataFields** (индекс), чтобы получить объект **MailMergeMappedDataField** . В этом примере возвращается имя поля источника данных в поле **pbFirstName** сопоставленные данные. В этом примере предполагается, что текущей публикации является публикацией слияния почты. Пустая строковое значение, возвращаемое для свойства **DataFieldName** указывает, что поле не связан с полем в источнике данных.
 

 

```
Sub MappedFieldName() 
 Dim strMappedDataField As String 
 With ActiveDocument.MailMerge.DataSource 
 strMappedDataField = .MappedDataFields(pbFirstName).DataFieldName 
 If strMappedDataField <> "" Then 
 MsgBox "The mapped data field 'FirstName' is mapped to " _ 
 &amp; .MappedDataFields(pbFirstName).DataFieldName &amp; "." 
 Else 
 MsgBox "The mapped data field 'FirstName' is not " &amp; _ 
 "mapped to any of the data fields in your " &amp; _ 
 "data source." 
 End If 
 End With 
End Sub
```


## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](mailmergemappeddatafield-application-property-publisher.md)|
|[DataFieldName](mailmergemappeddatafield-datafieldname-property-publisher.md)|
|[Index](mailmergemappeddatafield-index-property-publisher.md)|
|[Name](mailmergemappeddatafield-name-property-publisher.md)|
|[Родительский раздел](mailmergemappeddatafield-parent-property-publisher.md)|
|[Значение](mailmergemappeddatafield-value-property-publisher.md)|

