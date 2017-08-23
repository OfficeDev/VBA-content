---
title: "Метод Hyperlinks.Add (издатель)"
keywords: vbapb10.chm6881284
f1_keywords: vbapb10.chm6881284
ms.prod: publisher
api_name: Publisher.Hyperlinks.Add
ms.assetid: f5a8cc01-a571-623d-bfab-fe48e43a21b1
ms.date: 06/08/2017
ms.openlocfilehash: 88eecf826503710fc1f49ce8dd5879b087532238
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="hyperlinksadd-method-publisher"></a>Метод Hyperlinks.Add (издатель)

Добавляет новый объект **гиперссылки** определенной коллекции **гиперссылки** и возвращает новый объект **гиперссылки** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **Добавление** ( **_Текст_**, **_адрес_**, **_RelativePage_**, **_PageID_**, **_TextToDisplay_**)

 переменная _expression_A, представляющий объект **гиперссылки** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Text|Обязательное свойство.| **TextRange**| Объект **TextRange** . Диапазон текста для преобразования его в гиперссылку.|
|Адрес|Необязательный| **String**|Адрес новой гиперссылки. Если RelativePage **pbHlinkTargetTypeURL** (по умолчанию) или **pbHlinkTargetTypeEmail**, необходимо указать адрес или возникает ошибка.|
|RelativePage|Необязательный| **PbHlinkTargetType**| Тип гиперссылки для добавления.|
|PageID|Необязательный| **Длинный**|Идентификатор страницы конечной страницы для нового гиперссылки. Если RelativePage **pbHlinkTargetTypePageID**, должен быть указан PageID или возникает ошибка. Идентификатор страницы соответствует свойству **[PageID](hyperlink-pageid-property-publisher.md)** конечной страницы.|
|TextToDisplay|Необязательный| **String**|Отображаемый текст нового гиперссылки. Если указан, то **TextToDisplay** заменяет диапазон текста, указанный в аргументе **текста** .|

### <a name="return-value"></a>Возвращаемое значение

Hyperlink


## <a name="remarks"></a>Заметки

RelativePage может иметь одно из следующих констант **PbHlinkTargetType** . Значение по умолчанию — **pbHlinkTargetTypeURL**.



| **pbHlinkTargetTypeEmail**|| **pbHlinkTargetTypeFirstPage**|| **pbHlinkTargetTypeLastPage**|| **pbHlinkTargetTypeNextPage**|| **pbHlinkTargetTypePageID**|| **pbHlinkTargetTypePreviousPage**|| **pbHlinkTargetTypeURL**|

## <a name="example"></a>Пример

В следующем примере добавляется гиперссылки на фигуры одно и фигуры два по одному active публикации. Первая гиперссылка указывает на внешний веб-сайт, а второй указывает ссылку на четвертой странице в публикации. Фигура одно и фигуры двух должны быть текстовые поля и в публикации для работы этого примера необходимо быть по крайней мере четыре страницы.


```vb
Dim hypNew As Hyperlink 
Dim lngPageID As Long 
Dim strPage As String 
 
With ActiveDocument.Pages(1).Shapes(1).TextFrame 
 Set hypNew = .TextRange.Hyperlinks.Add(Text:=.TextRange, _ 
 Address:="http://www.tailspintoys.com/", _ 
 TextToDisplay:="Tailspin") 
End With 
 
lngPageID = ActiveDocument.Pages(4).PageID 
strPage = "Go to page " _ 
 &; Str(ActiveDocument.Pages(4).PageNumber) 
 
With ActiveDocument.Pages(1).Shapes(2).TextFrame 
 Set hypNew = .TextRange.Hyperlinks.Add(Text:=.TextRange, _ 
 RelativePage:=pbHlinkTargetTypePageID, _ 
 PageID:=lngPageID, _ 
 TextToDisplay:=strPage) 
End With
```


