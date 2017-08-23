---
title: "Свойство MailMergeDataSource.Included (издатель)"
keywords: vbapb10.chm6291465
f1_keywords: vbapb10.chm6291465
ms.prod: publisher
api_name: Publisher.MailMergeDataSource.Included
ms.assetid: 1cdac925-5fd6-e1d0-4612-0641e6057a7e
ms.date: 06/08/2017
ms.openlocfilehash: 87e4171aea28c669d315c70b338ce8c11d0d18f8
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatasourceincluded-property-publisher"></a>Свойство MailMergeDataSource.Included (издатель)

 **Значение true,** Если запись входит в слияния почты. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Включенные**

 переменная _expression_A, представляющий объект **вывода** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Используйте метод **[SetAllIncludedFlags](mailmergedatasource-setallincludedflags-method-publisher.md)** для задания состояния включена для всех записей слияния почты.


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


