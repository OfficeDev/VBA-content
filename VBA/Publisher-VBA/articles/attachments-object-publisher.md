---
title: "Объект вложения (издатель)"
keywords: vbapb10.chm9175039
f1_keywords: vbapb10.chm9175039
ms.prod: publisher
api_name: Publisher.Attachments
ms.assetid: 61957961-8c75-992f-159c-51412ed309ea
ms.date: 06/08/2017
ms.openlocfilehash: c1fdc5fb765c63c59ed27912ee09b3685205b9e4
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="attachments-object-publisher"></a>Объект вложения (издатель)

Коллекция объектов **[вложения](attachment-object-publisher.md)** , представляющий все вложения в сообщение электронной почты объединенных.
 


## <a name="remarks"></a>Заметки

Коллекции **вложения** соответствует списку вложения в окне " **вложения** " в диалоговом окне **Слияние по электронной почте** в интерфейсе пользователя Microsoft Publisher (в меню **файл** выберите команду **Отправить**сообщение, нажмите кнопку **Отправить слияния почты**и нажмите кнопку **Параметры**).
 

 
Чтобы добавить объект **вложения** в коллекцию **вложений** и таким образом добавить вложения в списке вложений в объединенные сообщение электронной почты, который требуется отправить, используйте метод **Attachments.Add** .
 

 
Чтобы удалить одного вложения из сообщения электронной почты объединения, используйте метод **Attachment.Delete** определенного объекта **вложения** , которое требуется удалить из коллекции **вложения** .
 

 
Чтобы удалить все вложения в объединенном сообщение электронной почты и таким образом, пустая коллекция **вложения** , используйте метод **Attachments.ClearAll** .
 

 
Свойство по умолчанию коллекции **вложения** — это свойство **Item** .
 

 

## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как использовать метод **Add** для добавления вложения в сообщение электронной почты merge. Макрос добавляет объект **вложения** , представляющий растрового изображения в коллекцию **вложений** активных документов. Также итерацию по коллекции **вложений** и печатает имя каждого вложения в окне **Интерпретация** .
 

 
Прежде чем запустить этот макрос, поместите файл с именем _image.bmp_ в корне диска C на вашем компьютере или измените имя и путь к файлу в макросе для указания на то, что необходимо присоединить.
 

 
Отправка сообщения электронной почты объединения, необходимо подключение к источнику данных, создание слияния почты и отправьте сообщение. Дополнительные сведения см в разделе объекта **[EmailMergeEnvelope](emailmergeenvelope-object-publisher.md)** .
 

 



```
Public Sub Attachments_Example() 
 
 Dim pubAttachments As Publisher.Attachments 
 Dim pubAttachment As Publisher.Attachment 
 Dim pubAttachment_Added As Publisher.Attachment 
 Dim pubMailMerge As Publisher.MailMerge 
 Dim pubEmailMergeEnvelope As Publisher.EmailMergeEnvelope 
 
 Set pubMailMerge = ThisDocument.MailMerge 
 Set pubEmailMergeEnvelope = pubMailMerge.EmailMergeEnvelope 
 Set pubAttachments = pubEmailMergeEnvelope.Attachemts 
 
 Set pubAttachment_Added = pubAttachments.Add("C:\image.bmp ") 
 
 For Each pubAttachment In pubAttachments 
 Debug.Print pubAttachment.Name 
 Next 
 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Добавление](attachments-add-method-publisher.md)|
|[ClearAll](attachments-clearall-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](attachments-application-property-publisher.md)|
|[Count](attachments-count-property-publisher.md)|
|[Элемент](attachments-item-property-publisher.md)|
|[Родительский раздел](attachments-parent-property-publisher.md)|

