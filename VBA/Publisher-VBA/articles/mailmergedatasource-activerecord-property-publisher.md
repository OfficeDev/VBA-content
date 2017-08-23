---
title: "Свойство MailMergeDataSource.ActiveRecord (издатель)"
keywords: vbapb10.chm6291459
f1_keywords: vbapb10.chm6291459
ms.prod: publisher
api_name: Publisher.MailMergeDataSource.ActiveRecord
ms.assetid: 0f092eb4-6e65-9235-83e2-a04b813b2390
ms.date: 06/08/2017
ms.openlocfilehash: d83e1efb2c86cd0614e77d363252270cb4f6e5ac
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatasourceactiverecord-property-publisher"></a>Свойство MailMergeDataSource.ActiveRecord (издатель)

Возвращает или задает **времени** , представляющий запись active слияния почты. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ActiveRecord**

 переменная _expression_A, представляющий объект **вывода** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="remarks"></a>Заметки

Номер активной записи — положение записи в результатах запроса, созданные средством параметры текущего запроса; Таким образом этот номер не обязательно положение записи в источнике данных.


## <a name="example"></a>Пример

В этом примере проверяет, что значение, указанное в поле «Индекс» является 10 символов (ПОЧТОВЫЙ индекс США плюс 4-значного локатор кода). Если он не установлен, исключены из слияния почты и помеченные комментарий.


```vb
Sub ValidateZip() 
 
 Dim intCount As Integer 
 
 On Error Resume Next 
 
 With ActiveDocument.MailMerge.DataSource 
 
 'Set the active record equal to the first included 
 'record in the data source 
 .ActiveRecord = 1 
 Do 
 intCount = intCount + 1 
 
 'Set the condition that the PostalCode field 
 'must be greater than or equal to ten digits 
 If Len(.DataFields.Item("PostalCode").Value) < 10 Then 
 
 'Exclude the record if the PostalCode field 
 'is less than ten digits 
 .Included = False 
 
 'Mark the record as containing an invalid address field 
 .InvalidAddress = True 
 
 'Specify the comment attached to the record explaining 
 'why the record was excluded from the mail merge 
 .InvalidComments = "The ZIP code for this record is " _ 
 &; "less than ten digits. It will be removed " _ 
 &; "from the mail merge process." 
 
 End If 
 
 'Move the record to the next record in the data source 
 .ActiveRecord = .ActiveRecord + 1 
 
 'End the loop when the counter variable 
 'equals the number of records in the data source 
 Loop Until intCount = .RecordCount 
 End With 
 
End Sub
```


