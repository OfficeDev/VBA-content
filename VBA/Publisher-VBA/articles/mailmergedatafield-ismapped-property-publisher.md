---
title: "Свойство MailMergeDataField.IsMapped (издатель)"
keywords: vbapb10.chm6422565
f1_keywords: vbapb10.chm6422565
ms.prod: publisher
api_name: Publisher.MailMergeDataField.IsMapped
ms.assetid: 4a053a2b-f6ca-37a7-4a1f-8690982188c2
ms.date: 06/08/2017
ms.openlocfilehash: 101c3a459fcfe5e2bd42de5c83bc1ad481a5848d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatafieldismapped-property-publisher"></a>Свойство MailMergeDataField.IsMapped (издатель)

Указывает, если родительский объект **MailMergeDataField** сопоставляется к полю получателя в источник данных (список получателей объединенный слияния почты). Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **IsMapped**

 переменная _expression_A, представляет собой объект- **MailMergeDataField** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Родительский объект **MailMergeDataField** должен представлять поля (столбца) для источника данных, который не является источником основных данных (сочетание всех подключенных источников данных). Свойство **IsMapped** недоступно для полей данных в источнике данных, представленного свойство **DataSource** объекта **слияния** активной объекта **Document** ( `ThisDocument.MailMerge.DataSource`).


