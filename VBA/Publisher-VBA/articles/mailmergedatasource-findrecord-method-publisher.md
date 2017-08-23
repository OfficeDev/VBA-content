---
title: "Метод MailMergeDataSource.FindRecord (издатель)"
keywords: vbapb10.chm6291480
f1_keywords: vbapb10.chm6291480
ms.prod: publisher
api_name: Publisher.MailMergeDataSource.FindRecord
ms.assetid: a4b37255-bdff-ac61-6d18-05a4fe008beb
ms.date: 06/08/2017
ms.openlocfilehash: dc445b50892eed9827e9d46c350d8d9687ccabd8
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatasourcefindrecord-method-publisher"></a>Метод MailMergeDataSource.FindRecord (издатель)

Осуществляет поиск содержимого источника данных указанного слияния почты для текста в отдельного поля. Возвращает значение **типа Boolean** , указывающее, будет ли найден искомый текст; **Значение true,** Если текст поиска найден.


## <a name="syntax"></a>Синтаксис

 _выражение_. **НайтиЗапись** ( **_FindText_**, **_поле_**)

 переменная _expression_A, представляющий объект **вывода** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|FindText|Обязательное свойство.| **String**|Текст для поиска.|
|Поле|Необязательный| **String**|Имя поля для поиска.|

### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="example"></a>Пример

В этом примере отображаются публикации слиянием для первой записи, в котором поля FirstName содержит Джо. Если запись найдена, номер записи хранится в переменной.


```vb
Sub FindDataSourceRecord() 
 Dim dsMain As MailMergeDataSource 
 Dim intRecord As Integer 
 
 'Makes the data in the data source records instead of the field codes 
 ActiveDocument.MailMerge.ViewMailMergeFieldCodes = False 
 
 Set dsMain = ActiveDocument.MailMerge.DataSource 
 
 If dsMain.FindRecord(FindText:="Joe", _ 
 Field:="FirstName") = True Then 
 intRecord = dsMain.ActiveRecord 
 End If 
 
End Sub
```


