---
title: "Метод Attachment.Delete (издатель)"
keywords: vbapb10.chm573441
f1_keywords: vbapb10.chm573441
ms.prod: publisher
api_name: Publisher.Attachment.Delete
ms.assetid: 935fa9e7-9d40-b820-e386-1a1960845da1
ms.date: 06/08/2017
ms.openlocfilehash: 6be662c026bbd5e96c970ff2c2bffa41d2b51852
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="attachmentdelete-method-publisher"></a>Метод Attachment.Delete (издатель)

Удаляет объект **вложения** из коллекции **вложения** сообщения электронной почты merge.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Удаление**

 переменная _expression_A, представляющий объект **вложения** .


## <a name="remarks"></a>Заметки

Метод **Delete** выполняет операцию нельзя обратить на коллекцию **вложений** . Он вызывает **IUnknown.Release** на коллекцию ссылку на объект **вложения** . Если у вас есть другой ссылку на вложение, по-прежнему доступ к его свойствам и методам, но его можно снова никогда не связать с любой коллекции, так как метод **[Add](attachments-add-method-publisher.md)** всегда создает новый объект. Ключевое слово **задать** Установка ссылочной переменной значение **Nothing** или другой вложения.

Окончательной версии объект **вложения** выполняется при назначении ссылочной переменной значение **Nothing**или при вызове, **Удаление**, если у вас есть не Справочник по. На этом этапе объект удаляется из памяти. При попытке получить доступ к объекту выпущенная возвращает ошибку объектов данных совместной работы Microsoft **CdoE_INVALID_OBJECT**.

При удалении элемент коллекции коллекции немедленно обновляется, что означает, что его свойство **Count** уменьшается на единицу и его члены являются переиндексации. Для доступа к элемент, который ранее, а затем удалить элемент в коллекции, необходимо использовать его нового значения индекса.

Чтобы удалить все вложения в текущем merge сообщение электронной почты, используйте метод **[ClearAll](attachments-clearall-method-publisher.md)** коллекции **вложения** .


## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как удалять вложения в сообщение электронной почты merge. Код удаляет вложения с первой позиции индекса в коллекции **вложений** и распечатывает имя удаленного вложения и номер текущего вложения в сообщение в окне **Интерпретация** .

Перед запуском этого кода убедитесь, что имеется по крайней мере один вложений в текущем merge сообщение электронной почты.




```vb
Public Sub Delete_Example() 
 
 Dim pubAttachments As Publisher.Attachments 
 Dim pubAttachment As Publisher.Attachment 
 
 Dim pubMailMerge As Publisher.MailMerge 
 Dim pubEmailMergeEnvelope As Publisher.EmailMergeEnvelope 
 
 Set pubMailMerge = ThisDocument.MailMerge 
 Set pubEmailMergeEnvelope = pubMailMerge.EmailMergeEnvelope 
 Set pubAttachments = pubEmailMergeEnvelope.Attachemts 
 
 Set pubAttachment = pubAttachments(1) 
 Debug.Print pubAttachments.Count 
 Debug.Print pubAttachment.Name 
 
 pubAttachment.Delete 
 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект вложения](attachment-object-publisher.md)

