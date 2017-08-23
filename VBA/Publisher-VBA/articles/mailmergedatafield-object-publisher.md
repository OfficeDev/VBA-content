---
title: "Объект MailMergeDataField (издатель)"
keywords: vbapb10.chm6488063
f1_keywords: vbapb10.chm6488063
ms.prod: publisher
api_name: Publisher.MailMergeDataField
ms.assetid: 46768b72-482c-06c5-5e77-27a95109f610
ms.date: 06/08/2017
ms.openlocfilehash: 0dea9ba17540b42cb6c0aaf31235fa4d1195e301
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatafield-object-publisher"></a>Объект MailMergeDataField (издатель)

Представляет поле одного объединения в источнике данных. Объект **MailMergeDataField** является элементом коллекции **[MailMergeDataFields](mailmergedatafields-object-publisher.md)** . Коллекция **MailMergeDataFields** включает все поля данных для слияния почты и каталогов объединения данных источника (например, имя, адрес и Город).
 


## <a name="remarks"></a>Заметки

Нельзя добавлять поля в коллекцию **MailMergeDataFields** . Все поля данных в источнике данных, автоматически включается в коллекции **MailMergeDataFields** .
 

 

## <a name="example"></a>Пример

Используйте **[DataFields](mailmergedatasource-datafields-property-publisher.md)** (индекс), где индекс — это имя поля данных или порядковый номер, чтобы получить объект **MailMergeDataField** . Номер индекса представляет позицию поля данных в источнике данных. В этом примере извлекается имя первого поля и значения первой записи поля FirstName в источнике данных, подключенного к active публикации.
 

 

```
Sub GetDataFromSource() 
 With ActiveDocument.MailMerge.DataSource 
 MsgBox "Field Name: " &amp; .DataFields.Item(1).Name &amp; _ 
 "Value: " &amp; .DataFields.Item("FirstName").Value 
 End With 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[AddToRecipientFields](mailmergedatafield-addtorecipientfields-method-publisher.md)|
|[Вставка](mailmergedatafield-insert-method-publisher.md)|
|[MapToRecipientField](mailmergedatafield-maptorecipientfield-method-publisher.md)|
|[UnMapRecipientField](mailmergedatafield-unmaprecipientfield-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](mailmergedatafield-application-property-publisher.md)|
|[Создатель](mailmergedatafield-creator-property-publisher.md)|
|[FieldType](mailmergedatafield-fieldtype-property-publisher.md)|
|[Index](mailmergedatafield-index-property-publisher.md)|
|[IsMapped](mailmergedatafield-ismapped-property-publisher.md)|
|[MappedTo](mailmergedatafield-mappedto-property-publisher.md)|
|[Name](mailmergedatafield-name-property-publisher.md)|
|[Родительский раздел](mailmergedatafield-parent-property-publisher.md)|
|[Значение](mailmergedatafield-value-property-publisher.md)|

