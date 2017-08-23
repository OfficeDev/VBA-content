---
title: "Метод Attachments.ClearAll (издатель)"
keywords: vbapb10.chm569350
f1_keywords: vbapb10.chm569350
ms.prod: publisher
api_name: Publisher.Attachments.ClearAll
ms.assetid: ae4e4c60-56cb-f97b-06f4-bd0d2abac4ee
ms.date: 06/08/2017
ms.openlocfilehash: eb49f80a615f8f8b54ae6aa81539d887bfe2ed16
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="attachmentsclearall-method-publisher"></a>Метод Attachments.ClearAll (издатель)

Удаляет объекты (удаление) все **вложения** в родительской коллекции **вложения** сообщения электронной почты merge.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ClearAll**

 переменная _expression_A, представляющий коллекцию **вложений** .


## <a name="remarks"></a>Заметки

Чтобы очистить отдельного вложения, используйте метод **[Delete](attachment-delete-method-publisher.md)** определенного объекта **вложения**


## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как очистить все вложения в сообщение электронной почты merge. Код печатает номер текущего вложения в сообщение в окне **интерпретации** и затем удаляет все объекты **вложения** в коллекции.


```vb
Public Sub ClearAll_Example() 
 
 Dim pubAttachments As Publisher.Attachments 
 
 Dim pubMailMerge As Publisher.MailMerge 
 Dim pubEmailMergeEnvelope As Publisher.EmailMergeEnvelope 
 
 Set pubMailMerge = ThisDocument.MailMerge 
 Set pubEmailMergeEnvelope = pubMailMerge.EmailMergeEnvelope 
 Set pubAttachments = pubEmailMergeEnvelope.Attachemts 
 
 Debug.Print pubAttachments.Count 
 pubAttachments.ClearAll 
 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Коллекция вложений](attachments-object-publisher.md)

