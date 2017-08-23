---
title: "Событие Application.MailMergeGenerateBarcode (издатель)"
keywords: vbapb10.chm268435489
f1_keywords: vbapb10.chm268435489
ms.prod: publisher
api_name: Publisher.Application.MailMergeGenerateBarcode
ms.assetid: 5da4ec65-32b6-ea05-09ad-d2224eafee30
ms.date: 06/08/2017
ms.openlocfilehash: c1cddcee9385d0781d4c1d6d10438fab7cca115b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationmailmergegeneratebarcode-event-publisher"></a>Событие Application.MailMergeGenerateBarcode (издатель)

Происходит, когда Microsoft Publisher требует данные для создания штрих-кодов в публикации слияния почты, в частности при изменении списка получателей слияния почты.


## <a name="syntax"></a>Синтаксис

 _выражение_. **MailMergeGenerateBarcode** ( **_Doc_**, **_bstrString_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Doc|Обязательное свойство.| **Документ**|Текущей публикации.|
|bstrString|Обязательное свойство.| **String**|Выходной параметр. Строковое представление штрих-кода.|

## <a name="remarks"></a>Заметки

Надстройки сторонних производителей, которые проверяют адреса слияния почты событие **MailMergeGenerateBarcode** можно использовать для прослушивания действия пользователя, запрашивающего создаваться, штрих-коды. В этом случае, если надстройка получает уведомление, что событие **MailMergeGenerateBarcode** , и если активный документ подключен к источнику данных, надстройки можно использовать ** [MailMergeDataSource.ActiveRecord](mailmergedatasource-activerecord-property-publisher.md)** свойства, чтобы определить записи, для которой создается штрих-кода. Если активный документ не подключен к источнику данных, надстройка использует адрес текст напрямую.

Если надстройка можно использовать в текст адреса непосредственно, возвращает строковое представление штрих-кода для параметра output bstrString. Если надстройка нельзя использовать в текст адреса напрямую, возвращает пустую строку.

Чтобы разрешить запуск события **MailMergeGenerateBarcode** , должен обрабатывать событие **[MailMergeInsertBarcode](application-mailmergeinsertbarcode-event-publisher.md)** в коде и надстройки задать параметр OkToInsert, передаваемый событие значение **True**. 

Дополнительные сведения об использовании событий с помощью объекта **приложения** [С помощью событий объекта](using-events-with-the-application-object-publisher.md)см.


## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как обрабатывать события **MailMergeGenerateBarcode** . Возвращает строку, представляющую штрих-кода для активной записи. Обратите внимание на то, что переменная _indexNumberOfBarcodeColumn_ представляет индекс столбца в источнике данных, в котором приведены штрих-коды. В этом коде предполагается, что текущей публикации подключен к источнику данных.


```vb
Private Sub pubApplication_MailMergeGenerateBarcode(ByVal Doc As Document, bstrString As String) 
 bstrString = pubApplication.ActiveDocument.MailMerge.DataSource.DataFields.Item(indexNumberOfBarcodeColumn).Value 
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

