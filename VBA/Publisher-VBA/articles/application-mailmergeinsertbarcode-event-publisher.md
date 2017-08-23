---
title: "Событие Application.MailMergeInsertBarcode (издатель)"
keywords: vbapb10.chm268435481
f1_keywords: vbapb10.chm268435481
ms.prod: publisher
api_name: Publisher.Application.MailMergeInsertBarcode
ms.assetid: 6b901953-eaff-0189-1d33-678e935a2f7e
ms.date: 06/08/2017
ms.openlocfilehash: d87a00146a3a7ae752890e74b1e74a3918a5e911
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationmailmergeinsertbarcode-event-publisher"></a>Событие Application.MailMergeInsertBarcode (издатель)

Происходит, когда пользователь выполняет команду для вставки почтовых штрих-кодов в публикацию слияния почты в Microsoft Publisher пользовательского интерфейса (UI), либо программными средствами.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MailMergeInsertBarcode** ( **_Doc_**, **_OkToInsert_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Doc|Обязательное свойство.| **Документ**|Текущей публикации.|
|OkToInsert|Обязательное свойство.| **Boolean**|Выходной параметр.  **Значение true,** Если это хорошо для вставки штрих-коды.|

## <a name="remarks"></a>Заметки

Метод **[InsertBarcode](textrange-insertbarcode-method-publisher.md)** можно использовать для вставки штрих-кодов в публикации слияния.

Надстройки сторонних производителей, которые проверяют адреса слияния почты событие **MailMergeInsertBarcode** можно использовать для прослушивания действия пользователя, запрашивающего вставить, штрих-коды. В этом случае при надстройки получает уведомление, событие **MailMergeInsertBarcode** , проверяет допустимость адреса в списке слияния почты и если адресах являются допустимыми, она пытается создать штрих-коды. Если попытка успешна, надстройки должен возвращать **значение True** для параметра OkToInsert. В случае сбоя установки надстройки должен возвращать **значение False**.

Фактический штрих-код предоставляется издателю **[MailMergeGenerateBarcode](application-mailmergegeneratebarcode-event-publisher.md)** события.

Событие **MailMergeInsertBarcode** также происходит, когда пользователь щелкает **Добавить почтовый штрих-код** в области задач для **слияния почты** и **Объединение в каталог** или **Добавить почтовых штрих-кодов** в области задач **Задачи Publisher** в пользовательском Интерфейсе Publisher. Перед пользователь может щелкнуть любой из этих команд пользовательского интерфейса, вам необходимо сначала сделать их доступными путем установки свойства **[InsertBarcodeVisible](application-insertbarcodevisible-property-publisher.md)** значение **True**. 

Дополнительные сведения об использовании событий с помощью объекта **приложения** [С помощью событий объекта](using-events-with-the-application-object-publisher.md)см.


## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как обрабатывать события **MailMergeInsertBarcode** . Отображается сообщение с вопросом, следует ли продолжить Вставка штрих-коды.


```vb
Private Sub pubApplication_MailMergeInsertBarcode(ByVal Doc As Document, OkToInsert As Boolean) 
 
 Dim intResponse As Integer 
 
 intResponse = MsgBox("Proceed to insert barcodes?", vbYesNo) 
 
 If intResponse = vbYes Then OkToInsert = True 
 
End Sub
```

Для чтобы произошло это событие необходимо включить следующую строку кода в разделе **Общие описаний** модуля.




```vb
Public WithEvents pubApplication As Application
```

Затем выполните следующую процедуру инициализации.




```vb
Public Sub Initialize_pubApplication() 
 Set pubApplication = Publisher.Application 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

