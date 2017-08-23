---
title: "Метод MailMergeDataSource.EditRecord (издатель)"
keywords: vbapb10.chm6291504
f1_keywords: vbapb10.chm6291504
ms.prod: publisher
api_name: Publisher.MailMergeDataSource.EditRecord
ms.assetid: 1fa31b25-b00a-9478-b341-094c2cdb2d9e
ms.date: 06/08/2017
ms.openlocfilehash: a5ab3b471a47a1b5d9b245ddbc0ffcec7c49b843
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatasourceeditrecord-method-publisher"></a>Метод MailMergeDataSource.EditRecord (издатель)

Изменение одного из полей данных в одной из записей в источник данных (объединенный слияния почты списка получателей).


## <a name="syntax"></a>Синтаксис

 _выражение_. **ИзменитьЗапись** ( **_lRec_**, **_varField_** **_значение_**)

 переменная _expression_A, представляющий объект **вывода** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|lRec|Обязательное свойство.| **Длинный**|Идентификатор записи, которую требуется изменить.|
|varField|Обязательное свойство.| **Variant**|Поля данных (столбца), содержащую значение, которое требуется изменить.|
|Значение|Обязательное свойство.| **Variant**|Значение, которое нужно изменить.|

## <a name="remarks"></a>Заметки

Метод **ИзменитьЗапись** исправьте источника данных, который находится в ошибки, такие как устаревшие адрес получателя.

Метод **ИзменитьЗапись** не вносите никаких изменений в отдельные источники данных, составляющих основного источника данных.


## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как использовать метод **ИзменитьЗапись** изменение значения одного столбца в определенной записи в источник данных (объединенный слияния почты списка получателей).

Прежде чем запустить этот макрос, замените _recordID_ номер индекса записи в источнике данных, которую требуется изменить; Замените _fieldname_ имя поля (столбца) в записи, которую требуется изменить; и замените _ключевое значение_ нового значения, которые необходимо задать для поля.




```vb
Public Sub EditRecord_Example() 
 
 Dim pubMailMergeDataSource As Publisher.MailMergeDataSource 
 
 Set pubMailMergeDataSource = ThisDocument.MailMerge.DataSource 
 
 pubMailMergeDataSource.EditRecord recordID, "fieldname", "value" 
 
End Sub
```


