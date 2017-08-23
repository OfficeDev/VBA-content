---
title: "Свойство MailMergeDataSource.InvalidAddress (издатель)"
keywords: vbapb10.chm6291472
f1_keywords: vbapb10.chm6291472
ms.prod: publisher
api_name: Publisher.MailMergeDataSource.InvalidAddress
ms.assetid: c1857edc-260b-c9c2-8624-d6628e0733c4
ms.date: 06/08/2017
ms.openlocfilehash: c67e8b1517bb18ab6c68ffb2cf77d393d1de2434
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatasourceinvalidaddress-property-publisher"></a>Свойство MailMergeDataSource.InvalidAddress (издатель)

 **Значение true,** чтобы отметить записи в источнике данных, если оно содержит недопустимые данные. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **InvalidAddress**

 переменная _expression_A, представляющий объект **вывода** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Используйте метод **[SetAllErrorFlags](mailmergedatasource-setallerrorflags-method-publisher.md)** для задания свойств и **InvalidAddress** и **[InvalidComments](mailmergedatasource-invalidcomments-property-publisher.md)** для всех записей в источнике данных.


## <a name="example"></a>Пример

В этом примере выполняется поиск записей, убедитесь, что длина поля PostalCode для каждой записи срок, по крайней мере пяти цифр. Если он не установлен, запись исключены из слияния почты и помечается как недопустимый.


```vb
Sub ExcludeRecords() 
 Dim intRecord As Integer 
 With ActiveDocument.MailMerge 
 For intRecord = 1 To .DataSource.RecordCount 
 .DataSource.ActiveRecord = intRecord 
 If Len(.DataSource.DataFields("PostalCode").Value) < 5 Then 
 With .DataSource 
 .Included = False 
 .InvalidAddress = True 
 .InvalidComments = "This record is removed " &; _ 
 "from the mail merge because its postal code" &; _ 
 "has less than five digits." 
 End With 
 End If 
 Next 
 End With 
End Sub
```


