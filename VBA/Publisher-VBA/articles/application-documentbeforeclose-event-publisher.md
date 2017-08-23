---
title: "Событие Application.DocumentBeforeClose (издатель)"
keywords: vbapb10.chm268435464
f1_keywords: vbapb10.chm268435464
ms.prod: publisher
api_name: Publisher.Application.DocumentBeforeClose
ms.assetid: d3ca4397-4df3-dc77-b758-d47e0bf13fe5
ms.date: 06/08/2017
ms.openlocfilehash: c8745654781149703040c703af26ab453be0c3a9
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationdocumentbeforeclose-event-publisher"></a>Событие Application.DocumentBeforeClose (издатель)

Происходит непосредственно перед закрытием любого открытого документа.


## <a name="syntax"></a>Синтаксис

 _выражение_. **DocumentBeforeClose** ( **_Doc_**, **_Отменить_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Doc|Обязательное свойство.| **Документ**|Документ, который был закрыт.|
|Cancel|Обязательное свойство.| **Boolean**| **Значение false,** при возникновении события. Если этот аргумент задает процедуру события значение **True**, документ не закрывается после завершения процедуры.|

## <a name="remarks"></a>Заметки

Для доступа к событий объекта **приложения** , объявите объектную переменную **приложения** в разделе Общие описаний модуля кода. Задайте переменную равно объект **приложения** , для которого требуется получить доступ к событиям. Сведения об использовании событий с помощью объекта Microsoft Publisher **приложения** [С помощью событий объекта](using-events-with-the-application-object-publisher.md)см.


## <a name="example"></a>Пример

В этом примере пользователю Да или нет ответа перед закрытием документа. Этот код для просмотра в этом примере работы, должны находиться в модуле класса и экземпляр класса необходимо правильно инициализировать, с помощью следующего ниже процедуры **SetPubApp** пример.


```vb
Private WithEvents PubApp As Application 
 
Sub SetPubApp() 
 Set PubApp = Publisher.Application 
End Sub 
 
Private Sub PubApp_DocumentBeforeClose(ByVal Doc As Document, Cancel As Boolean) 
 
 Dim intResponse As Integer 
 
 intResponse = MsgBox("Do you really want to close " _ 
 &; "the document?", vbYesNo) 
 
 If intResponse = vbNo Then Cancel = True 
 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

