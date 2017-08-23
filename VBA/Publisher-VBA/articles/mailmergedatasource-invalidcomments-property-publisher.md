---
title: "Свойство MailMergeDataSource.InvalidComments (издатель)"
keywords: vbapb10.chm6291473
f1_keywords: vbapb10.chm6291473
ms.prod: publisher
api_name: Publisher.MailMergeDataSource.InvalidComments
ms.assetid: ee08b03a-57e2-d79c-ee9f-a6f9231c8d6b
ms.date: 06/08/2017
ms.openlocfilehash: ee686b52d2a7fb3fa5772bac3b84fd5ef47995ce
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatasourceinvalidcomments-property-publisher"></a>Свойство MailMergeDataSource.InvalidComments (издатель)

Если свойство **[InvalidAddress](mailmergedatasource-invalidaddress-property-publisher.md)** имеет **значение True**, данное свойство Возвращает или задает **строку** , которая описывает недопустимые данные в записи. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **InvalidComments**

 переменная _expression_A, представляющий объект **вывода** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="remarks"></a>Заметки

Используйте метод **[SetAllErrorFlags](mailmergedatasource-setallerrorflags-method-publisher.md)** для задания свойств и **[InvalidAddress](mailmergedatasource-invalidaddress-property-publisher.md)** и **InvalidComments** для всех записей в источнике данных.


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


