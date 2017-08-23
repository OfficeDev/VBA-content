---
title: "Свойство EmailMergeEnvelope.Cc (издатель)"
keywords: vbapb10.chm9043972
f1_keywords: vbapb10.chm9043972
ms.prod: publisher
api_name: Publisher.EmailMergeEnvelope.Cc
ms.assetid: d9e7704c-c45a-cf19-e0a8-8d55e1e82fc0
ms.date: 06/08/2017
ms.openlocfilehash: fc4b4589e4f767131e1039217a6ad0afbcaa2438
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="emailmergeenvelopecc-property-publisher"></a>Свойство EmailMergeEnvelope.Cc (издатель)

Получает или задает объект **MailMergeDataField** , который представляет источник данных поля (столбца), в котором приведены адреса электронной почты получателей, вы хотите получать скрытую копию (CC) сообщения электронной почты слиянием. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **«Копия»**

 переменная _expression_A, представляющий объект **EmailMergeEnvelope** .


### <a name="return-value"></a>Возвращаемое значение

MailMergeDataField


## <a name="remarks"></a>Заметки

Необходимо включить для определенных, присвойте правильные источника данных поле (решает, соответствующий электронной почты «копия») в свойстве **«копия»** . Можно использовать следующую строку кода, который возвращает значение свойства **Name** объекта **MailMergeDataField** , к которому **«копия»** назначается, убедитесь, что правильный назначения:


```vb
Debug.Print ThisDocument.MailMerge.EmailMergeEnvelope.Cc.Name
```

Пример того, как задать значение свойства **«копия»** приведены в разделе объект **[EmailMergeEnvelope](emailmergeenvelope-object-publisher.md)** .


