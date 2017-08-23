---
title: "Метод Attachments.Add (издатель)"
keywords: vbapb10.chm569349
f1_keywords: vbapb10.chm569349
ms.prod: publisher
api_name: Publisher.Attachments.Add
ms.assetid: dbf2eb67-5e28-a7e6-226f-feac9045186b
ms.date: 06/08/2017
ms.openlocfilehash: ee7be82171cd5a3e6ea9f65e2dc076f699b18eab
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="attachmentsadd-method-publisher"></a>Метод Attachments.Add (издатель)

Добавляет объект **вложения** в коллекцию **вложений** публикации Microsoft Publisher.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Добавление** ( **_Имя файла_**)

 переменная _expression_A, представляющий colleciton **вложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Имя файла|Обязательное свойство.| **String**|Имя файла вложения.|

### <a name="return-value"></a>Возвращаемое значение

Attachment


## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как добавить вложения в сообщение в слияния почты. Код добавляет вложения в сообщение электронной почты, а затем печать номер текущего вложения в сообщение в окне **Интерпретация** .

Вложения в данном примере — это файл изображения в корне диска С. Перед выполнением кода, замените « _C:\image.jpg_» путь и имя файла на компьютере, который нужно добавить в качестве вложения электронной почты.

Перед созданием слияния почты необходимо использовать метод **[OpenDataSource](mailmerge-opendatasource-method-publisher.md)** объекта **[слияния](mailmerge-object-publisher.md)** для подключения к источнику данных активных документов. Чтобы выполнить слияние, используйте метод **[Execute](findreplace-execute-method-publisher.md)** объекта **слияния** . Пример того, как для подключения к источнику данных и создания слияния почты приведены в разделе объект **[EmailMergeEnvelope](emailmergeenvelope-object-publisher.md)** .




```vb
Public Sub Add_Example() 
 
 Dim pubAttachment As Publisher.Attachment 
 
 Set pubAttachment = ThisDocument.MailMerge.EmailMergeEnvelope.Attachemts.Add("C:\image.jpg") 
 Debug.Print ThisDocument.MailMerge.EmailMergeEnvelope.Attachemts.Count 
 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Коллекция вложений](attachments-object-publisher.md)

