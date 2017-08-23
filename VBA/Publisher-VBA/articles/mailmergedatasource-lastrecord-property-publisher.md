---
title: "Свойство MailMergeDataSource.LastRecord (издатель)"
keywords: vbapb10.chm6291474
f1_keywords: vbapb10.chm6291474
ms.prod: publisher
api_name: Publisher.MailMergeDataSource.LastRecord
ms.assetid: c1d11d3e-5f6f-2729-081b-5727c75fbc8d
ms.date: 06/08/2017
ms.openlocfilehash: 5ef59197caf7d1e39ea89bc5af1ec76e1e712181
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatasourcelastrecord-property-publisher"></a>Свойство MailMergeDataSource.LastRecord (издатель)

Возвращает или задает типа **Long** , представляющее номер последней записи, объединенных в операции объединения слияния почты и каталогов. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **LastRecord**

 переменная _expression_A, представляющий объект **вывода** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="example"></a>Пример

В этом примере показана активной записи как первой записи для объединения и затем последней записи как пересылать записей две записи в источнике данных. В этом примере предполагает, что активная публикация публикации слияния.


```vb
Sub RecordOne() 
 With ActiveDocument.MailMerge 
 .DataSource.FirstRecord = .DataSource.ActiveRecord 
 .DataSource.LastRecord = .DataSource.ActiveRecord + 2 
 .Execute Pause:=True 
 End With 
End Sub
```


