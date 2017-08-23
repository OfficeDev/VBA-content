---
title: "Свойство MailMergeDataSource.FirstRecord (издатель)"
keywords: vbapb10.chm6291464
f1_keywords: vbapb10.chm6291464
ms.prod: publisher
api_name: Publisher.MailMergeDataSource.FirstRecord
ms.assetid: e6eefea9-b353-27ff-d8e4-dc135c0c4665
ms.date: 06/08/2017
ms.openlocfilehash: e39b135c21ff2df52f84034b448abcc5328acb93
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatasourcefirstrecord-property-publisher"></a>Свойство MailMergeDataSource.FirstRecord (издатель)

Возвращает или задает типа **Long** , представляющее номер первой записи, объединенных в операции объединения слияния почты и каталогов. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **FirstRecord**

 переменная _expression_A, представляющий объект **вывода** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="example"></a>Пример

В этом примере задается активной записи первой записи для объединения, а затем объединяет три записи к концу с записью вперед две записи в источнике данных. В этом примере предполагает активная публикация документа слияния почты.


```vb
Sub RecordOne() 
 With ActiveDocument.MailMerge 
 .DataSource.FirstRecord = .DataSource.ActiveRecord 
 .DataSource.LastRecord = .DataSource.ActiveRecord + 2 
 .Execute Pause:=True 
 End With 
End Sub
```


