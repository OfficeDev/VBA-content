---
title: "Свойство MailMergeDataSource.DataFields (издатель)"
keywords: vbapb10.chm6291461
f1_keywords: vbapb10.chm6291461
ms.prod: publisher
api_name: Publisher.MailMergeDataSource.DataFields
ms.assetid: 820af882-d54c-a205-2925-e7110fc0c02b
ms.date: 06/08/2017
ms.openlocfilehash: c67cb066a0c44af0878748a2434eae740bb99df5
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatasourcedatafields-property-publisher"></a>Свойство MailMergeDataSource.DataFields (издатель)

Возвращает коллекцию **[MailMergeDataFields](mailmergedatafields-object-publisher.md)** , который представляет поля в указанном источнике данных.


## <a name="syntax"></a>Синтаксис

 _выражение_. **DataFields**

 переменная _expression_A, представляющий объект **вывода** .


### <a name="return-value"></a>Возвращаемое значение

MailMergeDataFields


## <a name="example"></a>Пример

В этом примере отображает значение поля FirstName и LastName из активной записи в источнике данных, подключенного к active публикации.


```vb
Sub ShowNameForActiveRecord() 
 Dim mdfFirst As MailMergeDataField 
 Dim mdfLast As MailMergeDataField 
 
 With ActiveDocument.MailMerge.DataSource 
 Set mdfFirst = .DataFields.Item("FirstName") 
 Set mdfLast = .DataFields.Item("LastName") 
 MsgBox "The active record in the attached " &; _ 
 vbLf &; "data source is : " &; _ 
 mdfFirst.Value &; " " &; _ 
 mdfLast.Value 
 End With 
End Sub
```


