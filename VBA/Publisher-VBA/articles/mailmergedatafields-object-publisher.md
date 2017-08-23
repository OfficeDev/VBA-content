---
title: "Объект MailMergeDataFields (издатель)"
keywords: vbapb10.chm6422527
f1_keywords: vbapb10.chm6422527
ms.prod: publisher
api_name: Publisher.MailMergeDataFields
ms.assetid: 44ae8a3c-b8a8-fc57-9d02-d71dcffc21ef
ms.date: 06/08/2017
ms.openlocfilehash: 54c8ba3306e3ff61b8ead4cb284c34eb408efc7e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatafields-object-publisher"></a>Объект MailMergeDataFields (издатель)

Коллекция объектов **[MailMergeDataField](mailmergedatafield-object-publisher.md)** , представляющих полей данных в источнике данных слияния слияния почты и каталогов.
 


## <a name="remarks"></a>Заметки

Нельзя добавлять поля в коллекцию **MailMergeDataFields** . При добавлении поля данных в источнике данных, это поле автоматически включается в коллекции **MailMergeDataFields** .
 

 

## <a name="example"></a>Пример

Свойство **[DataFields](mailmergedatasource-datafields-property-publisher.md)** используется для возврата коллекции **MailMergeDataFields** .
 

 

 

 
Следующий пример отображает имена полей в источнике данных, подключенного к active публикации.
 

 



```
Sub ShowFieldNames() 
 Dim intCount As Integer 
 With ActiveDocument.MailMerge.DataSource.DataFields 
 For intCount = 1 To .Count 
 MsgBox .Item(intCount).Name 
 Next 
 End With 
End Sub
```

Используйте **DataFields** (индекс), где индекс — это имя поля данных или номер индекса, чтобы получить объект **MailMergeDataField** . Номер индекса представляет позицию поля данных в источнике данных. В этом примере извлекается имя первого поля и значения первой записи поля FirstName в источнике данных, подключенного к active публикации.
 

 



```
Sub GetDataFromSource() 
 With ActiveDocument.MailMerge.DataSource.DataFields 
 MsgBox "First field name: " &amp; .Item(1).Name &amp; vbLf &amp; _ 
 "Value of the first record of the FirstName field: " &amp; _ 
 .Item("FirstName").Value 
 End With 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Элемент](mailmergedatafields-item-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](mailmergedatafields-application-property-publisher.md)|
|[Count](mailmergedatafields-count-property-publisher.md)|
|[Создатель](mailmergedatafields-creator-property-publisher.md)|
|[Родительский раздел](mailmergedatafields-parent-property-publisher.md)|

