---
title: "Свойство MailMergeDataSource.RecordCount (издатель)"
keywords: vbapb10.chm6291477
f1_keywords: vbapb10.chm6291477
ms.prod: publisher
api_name: Publisher.MailMergeDataSource.RecordCount
ms.assetid: 56b929bf-9b7f-dd83-98b7-35bf96028732
ms.date: 06/08/2017
ms.openlocfilehash: 2c2037b6678bbd2721d2d28984f34e4026a8110d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatasourcerecordcount-property-publisher"></a>Свойство MailMergeDataSource.RecordCount (издатель)

Возвращает значение типа **Long** , представляющее количество записей в источнике данных. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **RecordCount**

 переменная _expression_A, представляющий объект **вывода** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="example"></a>Пример

В этом примере выполняется проверка индексы в источник данных для пяти цифр. Если длина ПОЧТОВЫЙ индекс составляет менее пяти цифр, записи исключается из процесс слияния почты. В этом примере предполагается, что почтовые индексы являются ПОЧТОВЫЕ индексы США. Можно изменить в этом примере для поиска индексы, содержащие код локатор четырехзначное, добавляется в конец ПОЧТОВЫЙ индекс и затем исключить все записи, не содержащих кода локатор.


```vb
Sub Validate 
 
 Dim intCount As Integer 
 With ActiveDocument.MailMerge.DataSource 
 
 'Set the active record equal to the first included record in the 
 'data source 
 .ActiveRecord = 1 
 Do 
 intCount = intCount + 1 
 
 'Set the condition that field six must be greater than or 
 'equal to five digits in length 
 If Len(.DataFields.Item(6).Value) < 5 Then 
 
 'Exclude the record if field six contains fewer than five digits 
 .Included = False 
 
 'Mark the record as containing an invalid address field 
 .InvalidAddress = True 
 
 'Specify the comment attached to the record explaining 
 'why the record was excluded from the mail merge 
 .InvalidComments = "The ZIP Code for this record has " _ 
 &; "fewer than five digits. It will be removed " _ 
 &; "from the mail merge process." 
 
 End If 
 
 'Move the record to the next record in the data source 
 .ActiveRecord = .ActiveRecord + 1 
 
 'End the loop when the counter variable 
 'equals the number of records in the data source 
 Loop Until intCount = .RecordCount 
 End With 

```


