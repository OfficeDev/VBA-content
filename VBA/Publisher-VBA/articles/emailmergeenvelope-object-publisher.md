---
title: "Объект EmailMergeEnvelope (издатель)"
keywords: vbapb10.chm9109503
f1_keywords: vbapb10.chm9109503
ms.prod: publisher
api_name: Publisher.EmailMergeEnvelope
ms.assetid: 555dd80e-bac2-96dd-4256-ad1b8006da0f
ms.date: 06/08/2017
ms.openlocfilehash: b76642d76f88a45a44ec7bc0806ad3de0c194e5d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="emailmergeenvelope-object-publisher"></a>Объект EmailMergeEnvelope (издатель)

Представляет контейнер электронной почты (конверт), в котором размещается документ Microsoft Publisher, объединенные в слияния электронной почты.
 


## <a name="remarks"></a>Заметки

Свойства объекта **EmailMergeEnvelope** соответствуют сочетание необходимые и необязательные параметры в диалоговом окне **Слияние по электронной почте** в пользовательском интерфейсе Publisher (в меню **файл** выберите команду **Отправить**сообщение, нажмите кнопку **Отправить слияния почты**и нажмите кнопку **Параметры**). 
 

 
Прежде чем использовать метод **Execute** объекта **[слияния](mailmerge-object-publisher.md)** для отправки объединенных электронной почты, необходимо указать значение для свойства **для** объекта **EmailMergeEnvelope** или Publisher возвращает ошибку.
 

 

## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как назначить некоторые свойства объекта **EmailMergeEnvelope** , который представляет слияния почты и отправьте итогового сообщения электронной почты, приглашения. Макрос подключается к источнику данных, присваивает значения свойства **к** и **Subject** объекта **EmailMergeEnvelope** и добавляет текстовое поле, содержащее полей слияния и некоторые дополнительный текст в сообщение электронной почты. Затем он использует метод **Execute** объекта **слияния** для выполнения объединения и отправки сообщения электронной почты.
 

 
В этом примере указанный источник данных является простой вкладку deliimited текстовый файл, содержащий три столбца с заголовками «First», «Последняя» и «Адрес электронной почты» соответственно.
 

 
Перед выполнением кода, создайте текстовый файл, добавьте один или несколько строк данных, имя файла DataSource.txt и сохраните его на диск. Затем добавьте путь файла кода, заменив _PathToFile_ переменную path.
 

 
При выполнении кода в этом примере показан более чем один раз, который будет возникли ошибки, так как Publisher подключается к источнику данных каждый раз, когда выполняется код, приведшего к публикации, подключенная к нескольким источникам данных. При наличии нескольких подключений к источнику данных, Microsoft Publisher вставляет дополнительный столбец в источнике данных master (комбинированные) слияния почты для указания конкретного источника данных для каждой записи. В результате Publisher эффективно изменяется номер индекса все столбцы источника данных, создание индексов, используемых в этом примере кода (например, _MailMergeField1_ ) неправильные.
 

 



```
Public Sub EmailMergeEnvelope_Example() 
 
 Dim pubShape As Publisher.Shape 
 Dim pubMailMerge As Publisher.MailMerge 
 
 'Connect to the data source. 
 Set pubMailMerge = ThisDocument.MailMerge 
 pubMailMerge.OpenDataSource "PathToFile \DataSource.txt" 
 
 'Assign "E-mail Address" to the To field of the e-mail message. 
 pubMailMerge.EmailMergeEnvelope.To = pubMailMerge.DataSource.DataFields.Item(3) 
 
 'Add text to the Subject field of the e-mail message. 
 pubMailMerge.EmailMergeEnvelope.Subject = "Invitation" 
 
 'Insert two merge fields and some additional text in a text box in the body of the message. 
 Set pubShape = ThisDocument.Pages(1).Shapes.AddTextbox(pbTextOrientationHorizontal, 100, 100, 200, 100) 
 pubShape.TextFrame.TextRange.Text = "Dear " 
 pubShape.TextFrame.TextRange.InsertMailMergeField 1 
 pubShape.TextFrame.TextRange.InsertAfter " " 
 pubShape.TextFrame.TextRange.InsertMailMergeField 2 
 pubShape.TextFrame.TextRange.InsertAfter ": " 
 pubShape.TextFrame.TextRange.InsertAfter "You are invited!" 
 
 'Perform the merge. 
 pubMailMerge.Execute True, pbSendEmail 
 
 'Display a reminder 
 MsgBox "If your e-mail client is not already open, remember to open it and send the e-mail messages that are in the outbox." 
 
End Sub
```


## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](emailmergeenvelope-application-property-publisher.md)|
|[Attachemts](emailmergeenvelope-attachemts-property-publisher.md)|
|[Скрытой копии](emailmergeenvelope-bcc-property-publisher.md)|
|[«Копия»](emailmergeenvelope-cc-property-publisher.md)|
|[Родительский раздел](emailmergeenvelope-parent-property-publisher.md)|
|[Приоритет](emailmergeenvelope-priority-property-publisher.md)|
|[Subject](emailmergeenvelope-subject-property-publisher.md)|
|[Для](emailmergeenvelope-to-property-publisher.md)|

