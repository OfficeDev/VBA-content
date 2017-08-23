---
title: "Метод Document.SaveAs (издатель)"
keywords: vbapb10.chm196696
f1_keywords: vbapb10.chm196696
ms.prod: publisher
api_name: Publisher.Document.SaveAs
ms.assetid: ba8b85d7-8ca9-dcf5-12b4-4cabced743e6
ms.date: 06/08/2017
ms.openlocfilehash: bb567cb1148b4f50326d8c62ad5d539f8ed22650
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentsaveas-method-publisher"></a>Метод Document.SaveAs (издатель)

Сохранение указанной публикации новое имя или формат.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Сохранить как** ( **_Имя файла_**, **_Формат_** **_AddToRecentFiles_**)

 переменная _expression_A, представляющий объект **Document** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Имя файла|Необязательный| **Variant**|Имя для публикации. Значение по умолчанию — это имя текущего папка и файл. Если не был сохранен публикации, имя по умолчанию используется для примера, Publication1.pub. Если публикация с указанным именем уже существует, публикации перезаписывается без предварительного предупреждения пользователя.|
|Формат|Необязательный| **PbFileFormat**|Формат, в котором будет сохранен публикации.|
|AddToRecentFiles|Необязательный| **Boolean**| **Значение true,** Чтобы добавить публикации в список недавно использовавшихся файлов в меню файл. Значение по умолчанию — **True**.|

## <a name="remarks"></a>Заметки

Параметр Format может иметь одно из **PbFileFormat** константы объявляются в библиотеке типов Microsoft Publisher и показаны в следующей таблице. Значение по умолчанию — **pbFilePublication**.



| **pbFileHTMLFiltered**|| **pbFilePublication**|| **pbFilePublicationHTML**|| **pbFilePublisher2000**|| **pbFilePublisher98**|| **pbFileRTF**|| **pbFileWebArchive**| Если недостаточно памяти или места на диске для сохранения файла, возникает ошибка.

При вызове метода **SaveAs** всегда выполняется сохранение на переднем плане независимо от того, включен ли фоновое сохранение.


## <a name="example"></a>Пример

В этом примере сохраняет активная публикация в виде файла Microsoft Publisher 2000.


```vb
ActiveDocument.SaveAs FileName:="ReportPub2000.pub", Format:=pbFilePublisher2000
```

В этом примере сохраняет active публикации в виде Test.rtf в форматированный текст (RTF).




```vb
ActiveDocument.SaveAs FileName:="Test.rtf", Format:=pbFileRTF
```

В этом примере сохраняет active веб-публикации в виде набора отфильтрованные HTML-страниц и вспомогательные файлы. Обратите внимание, что данное расширение имени файла .htm автоматически добавляется значение параметра Filename Если значение параметра формат **pbFileHTMLFiltered**.




```vb
With ActiveDocument 
 .SaveAs Filename:="CompanyContacts", Format:=pbFileHTMLFiltered 
End With
```


