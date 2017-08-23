---
title: "Событие Application.MailMergeDataSourceLoad (издатель)"
keywords: vbapb10.chm268435475
f1_keywords: vbapb10.chm268435475
ms.prod: publisher
api_name: Publisher.Application.MailMergeDataSourceLoad
ms.assetid: afca3a05-d6a6-15f1-8cbf-593777066757
ms.date: 06/08/2017
ms.openlocfilehash: 988326d6d6fd292a4b2069817c8d083fc214abef
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationmailmergedatasourceload-event-publisher"></a>Событие Application.MailMergeDataSourceLoad (издатель)

Происходит при загрузке источника данных для слияния почты.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MailMergeDataSourceLoad** ( **_Doc_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Doc|Обязательное свойство.| **Документ**|Основной документ слияния почты.|

## <a name="remarks"></a>Заметки

Для доступа к событий объекта **приложения** , объявите объектную переменную **приложения** в разделе Общие описаний модуля кода. Задайте переменную равно объект **приложения** , для которого требуется получить доступ к событиям. Сведения об использовании событий с помощью объекта Microsoft Publisher **приложения** [С помощью событий объекта](using-events-with-the-application-object-publisher.md)см.


## <a name="example"></a>Пример

В этом примере выводится сообщение с именем файла источника данных источника данных в начале загрузки.


```vb
Private Sub MailMergeApp_MailMergeDataSourceLoad(ByVal Doc As Document) 
 Dim strDSName As String 
 Dim intDSLength As Integer 
 Dim intDSStart As Integer 
 
 'Pull out of the Name property (which includes path and file name) 
 'only the file name using Visual Basic commands Len, InStrRev, and Right 
 intDSLength = Len(ActiveDocument.MailMerge.DataSource.Name) 
 intDSStart = InStrRev(ActiveDocument.MailMerge.DataSource.Name, "\") 
 intDSStart = intDSLength - intDSStart 
 strDSName = Right(ActiveDocument.MailMerge.DataSource.Name, intDSStart) 
 
 'Deliver a message to user when data source is loading 
 MsgBox "Your data source, " &; strDSName &; ", is now loading." 
End Sub
```

Чтобы произошло это событие необходимо поместить следующую строку кода в разделе Общие описаний модуля и выполнить следующую процедуру инициализации.




```vb
Private WithEvents MailMergeApp As Application 
 
Sub InitializeMailMergeApp() 
 Set MailMergeApp = Publisher.Application 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

