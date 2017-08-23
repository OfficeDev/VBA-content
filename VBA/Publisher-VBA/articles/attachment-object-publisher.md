---
title: "Объект Attachment (издатель)"
keywords: vbapb10.chm9240575
f1_keywords: vbapb10.chm9240575
ms.prod: publisher
api_name: Publisher.Attachment
ms.assetid: d617bdf6-b0ba-be0d-0f72-f729010636c1
ms.date: 06/08/2017
ms.openlocfilehash: 0adcf4190f4971169d5d82d82144627b33322a66
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="attachment-object-publisher"></a>Объект Attachment (издатель)

Представляет вложения в сообщение электронной почты слиянием.


## <a name="remarks"></a>Заметки

Объект **вложения** соответствует одному из вложений в списке вложений в поле **вложения** в диалоговом окне **Слияние по электронной почте** в интерфейсе пользователя Microsoft Publisher. (В меню **файл** выберите пункт **Отправить сообщение**, нажмите кнопку **Отправить слияния почты**и выберите пункт **Свойства**.)

Для удаления вложения из объединенной почты, используйте метод **удаления** объекта **вложения** .


## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как использовать метод **Add** для добавления вложения в сообщение электронной почты merge. Добавление **вложения** object, представляющий растрового изображения в коллекцию **вложений** активных документов.

Прежде чем запустить этот макрос, поместите файл с именем _image.bmp_ в корне диска C на вашем компьютере или измените имя и путь файла в макросе для указания на то, что необходимо присоединить.

Обратите внимание, что для отправки сообщения электронной почты объединения, вы должны подключения к источнику данных, создание слияния почты и отправьте сообщение. Дополнительные сведения см в разделе объекта **[EmailMergeEnvelope](http://msdn.microsoft.com/library/555dd80e-bac2-96dd-4256-ad1b8006da0f%28Office.15%29.aspx)** .




```
Public Sub Attachment_Example() 
 
 Dim pubAttachments As Publisher.Attachments 
 Dim pubAttachment As Publisher.Attachment 
 Dim pubMailMerge As Publisher.MailMerge 
 Dim pubEmailMergeEnvelope As Publisher.EmailMergeEnvelope 
 
 Set pubMailMerge = ThisDocument.MailMerge 
 Set pubEmailMergeEnvelope = pubMailMerge.EmailMergeEnvelope 
 Set pubAttachments = pubEmailMergeEnvelope.Attachemts 
 
 Set pubAttachment = pubAttachments.Add("C:\image.bmp ") 
 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/935fa9e7-9d40-b820-e386-1a1960845da1%28Office.15%29.aspx)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Name](http://msdn.microsoft.com/library/7539a5ac-427f-0dfe-dc31-47ef9436fd14%28Office.15%29.aspx)|

## <a name="see-also"></a>См. также


#### <a name="other-resources"></a>Другие ресурсы


[Члены объекта вложения](http://msdn.microsoft.com/library/594cf3eb-73d8-afa9-b598-ab68066dde8b%28Office.15%29.aspx)
