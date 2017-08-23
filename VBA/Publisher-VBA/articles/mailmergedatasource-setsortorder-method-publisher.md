---
title: "Метод MailMergeDataSource.SetSortOrder (издатель)"
keywords: vbapb10.chm6291489
f1_keywords: vbapb10.chm6291489
ms.prod: publisher
api_name: Publisher.MailMergeDataSource.SetSortOrder
ms.assetid: 0ecb5f77-2cd1-92c6-b7f2-bf709f015ba5
ms.date: 06/08/2017
ms.openlocfilehash: 279096245e4f3f352d549d2fe6f657b42bb51b64
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatasourcesetsortorder-method-publisher"></a>Метод MailMergeDataSource.SetSortOrder (издатель)

Задает порядок сортировки для слияния данных.


## <a name="syntax"></a>Синтаксис

 _выражение_. **SetSortOrder** ( **_SortField1_**, **_SortAscending1_**, **_SortField2_**, **_SortAscending2_**, **_SortField3_**, **_SortAscending3_**)

 переменная _expression_A, представляющий объект **вывода** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|SortField1|Необязательный| **String**|Первое поле, по которому выполняется сортировка данных слияния почты. Значение по умолчанию — пустая строка.|
|SortAscending1|Необязательный| **Boolean**| **Значение true** (по умолчанию) для выполнения сортировки по возрастанию на SortField1; **Значение false** для выполнения по убыванию сортировка.|
|SortField2|Необязательный| **String**|Второе поле, по которому выполняется сортировка данных слияния почты. Значение по умолчанию — пустая строка.|
|SortAscending2|Необязательный| **Boolean**| **Значение true** (по умолчанию) для выполнения сортировки по возрастанию на SortField2; **Значение false** для выполнения по убыванию сортировка.|
|SortField3|Необязательный| **String**|Третий поле, по которому выполняется сортировка данных слияния почты. Значение по умолчанию — пустая строка.|
|SortAscending3|Необязательный| **Boolean**| **Значение true** (по умолчанию) для выполнения сортировки по возрастанию на SortField3; **Значение false** для выполнения по убыванию сортировка.|

## <a name="example"></a>Пример

Следующий пример сначала сортировка данных слияния почты на почтовый индекс в убывающем порядке, затем на последний и имени в порядке возрастания.


```vb
ActiveDocument.MailMerge.DataSource.SetSortOrder _ 
 SortField1:="ZIPCode", SortAscending1:=False, _ 
 SortField2:="LastName", SortField3:="FirstName"
```


