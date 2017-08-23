---
title: "Свойство EmailMergeEnvelope.To (издатель)"
keywords: vbapb10.chm9043971
f1_keywords: vbapb10.chm9043971
ms.prod: publisher
api_name: Publisher.EmailMergeEnvelope.To
ms.assetid: c9c470e8-1411-fda9-becf-5c932e97d98f
ms.date: 06/08/2017
ms.openlocfilehash: 5ad4c184c1833f913fdc9a4dde77eb384d7d35c1
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="emailmergeenvelopeto-property-publisher"></a>Свойство EmailMergeEnvelope.To (издатель)

Получает или задает объект **MailMergeDataField** , который представляет источник данных поля (столбца), в котором приведены адреса электронной почты получателей сообщения электронной почты слиянием. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Чтобы**

 переменная _expression_A, представляющий объект **EmailMergeEnvelope** .


### <a name="return-value"></a>Возвращаемое значение

MailMergeDataField


## <a name="remarks"></a>Заметки

Необходимо включить для определенных назначить правильные источник данных поля (один, представляющий адреса электронной почты) свойство **для** . Можно использовать следующую строку кода, который возвращает значение свойства **Name** объекта **MailMergeDataField** , **к которому назначена, чтобы убедитесь, что правильный назначения** :


```vb
Debug.Print ThisDocument.MailMerge.EmailMergeEnvelope.To.Name
```

Пример того, как задать значение **для** свойства приведены в разделе объект **[EmailMergeEnvelope](emailmergeenvelope-object-publisher.md)** .


