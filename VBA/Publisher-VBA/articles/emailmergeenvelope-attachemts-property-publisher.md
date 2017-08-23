---
title: "Свойство EmailMergeEnvelope.Attachemts (издатель)"
keywords: vbapb10.chm9043975
f1_keywords: vbapb10.chm9043975
ms.prod: publisher
api_name: Publisher.EmailMergeEnvelope.Attachemts
ms.assetid: 53948bf7-2727-7b9c-a645-c9b954d5e023
ms.date: 06/08/2017
ms.openlocfilehash: 1ee4200046a43aaa093a2cb73e1af069022f9de9
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="emailmergeenvelopeattachemts-property-publisher"></a>Свойство EmailMergeEnvelope.Attachemts (издатель)

Получает список объединенной почты вложения сообщения. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Attachemts**

 переменная _expression_A, представляющий объект **EmailMergeEnvelope** .


### <a name="return-value"></a>Возвращаемое значение

Вложения


## <a name="remarks"></a>Заметки

Добавление вложения в сообщение электронной почты объединенных, используйте метод **[Add](attachments-add-method-publisher.md)** объекта **[вложения](attachment-object-publisher.md)** . Для удаления вложения, используйте ** [Attachment.Delete](attachment-delete-method-publisher.md)** метода. Чтобы удалить все вложения, используйте метод **[ClearAll](attachments-clearall-method-publisher.md)** коллекции **[вложения](attachments-object-publisher.md)** .


