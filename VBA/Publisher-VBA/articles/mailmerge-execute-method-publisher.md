---
title: "Метод MailMerge.Execute (издатель)"
keywords: vbapb10.chm6225940
f1_keywords: vbapb10.chm6225940
ms.prod: publisher
api_name: Publisher.MailMerge.Execute
ms.assetid: edcabcc5-f2ce-53ce-d422-0d6fcb5f8a33
ms.date: 06/08/2017
ms.openlocfilehash: 62ead1dcfb8a70fbe4bb0edeb5f465ab3bde6ebe
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergeexecute-method-publisher"></a>Метод MailMerge.Execute (издатель)

Выполняет указанной операции объединения слияния почты и каталогов. Возвращает объект **[Document](document-object-publisher.md)** , представляющий новой или существующей публикации, указан как назначения результаты объединения. Возвращает **значение Nothing** , если на принтере выполняется слияние.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Выполнение** ( **_Приостановить_**, **_назначения_**, **_имя файла_**)

 переменная _expression_A, представляет собой объект- **слияния** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Приостановка|Обязательное свойство.| **Boolean**| **Значение true,** приложение Microsoft Publisher Приостановка и отображения диалогового окна по устранению неполадок, если обнаружена ошибка объединения. **Значение false,** следует ли игнорировать ошибки во время слияния или объединения в каталог.|
|Конечный объект|Необязательный| **PbMailMergeDestination**|Назначение результаты объединения слияния почты и каталогов. Указание **pbSendToPrinter** для результаты объединения каталога в ошибку времени выполнения.|
|Имя файла|Необязательный| **String**|Имя файла публикации, к которому необходимо добавить результаты объединения каталога.|

### <a name="return-value"></a>Возвращаемое значение

Документ


## <a name="remarks"></a>Заметки

Назначение может иметь одно из следующих констант **PbMailMergeDestination** . Значение по умолчанию — **pbSendToPrinter**.



| **pbSendToPrinter**|| **pbMergeToNewPublication**|| **pbMergeToExistingPublication**|

## <a name="example"></a>Пример

В этом примере выполняется слияния почты, если активная публикация — это основной документ с источником данных.


```vb
Sub ExecuteMerge() 
 Dim mrgDocument As MailMerge 
 Set mrgDocument = ActiveDocument.MailMerge 
 If mrgDocument.DataSource.ConnectString <> "" Then 
 mrgDocument.Execute Pause:=False 
 End If 
End Sub
```


