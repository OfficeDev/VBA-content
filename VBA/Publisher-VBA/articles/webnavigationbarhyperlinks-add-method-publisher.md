---
title: "Метод WebNavigationBarHyperlinks.Add (издатель)"
keywords: vbapb10.chm8585220
f1_keywords: vbapb10.chm8585220
ms.prod: publisher
api_name: Publisher.WebNavigationBarHyperlinks.Add
ms.assetid: 6cd0c43a-fec1-c9b8-dc86-00e1cc314087
ms.date: 06/08/2017
ms.openlocfilehash: 03c7905ffdf8a1939268f415430215734ef6f25a
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webnavigationbarhyperlinksadd-method-publisher"></a>Метод WebNavigationBarHyperlinks.Add (издатель)

Добавляет новый объект **гиперссылки** определенной коллекции **WebNavigationBarHyperlinks** и возвращает новый объект **гиперссылки** . .


## <a name="syntax"></a>Синтаксис

 _выражение_. **Добавление** ( **_Адрес_**, **_RelativePage_**, **_PageID_**, **_TextToDisplay_**, **_индекса_**)

 переменная _expression_A, представляет собой объект- **WebNavigationBarHyperlinks** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Адрес|Необязательный| **String**|Адрес новой гиперссылки. Если RelativePage **pbHlinkTargetTypeURL** (по умолчанию) или **pbHlinkTargetTypeEmail**, необходимо указать адрес или возникает ошибка.|
|RelativePage|Необязательный| **PbHlinkTargetType**|Тип гиперссылки для добавления.|
|PageID|Необязательный| **Длинный**|Идентификатор страницы конечной страницы для нового гиперссылки. Если RelativePage **pbHlinkTargetTypePageID**, должен быть указан PageID или возникает ошибка. Идентификатор страницы соответствует свойству [PageID](page-pageid-property-publisher.md) конечной страницы.|
|TextToDisplay|Необязательный| **String**|Отображаемый текст нового гиперссылки. |
|Индекс|Необязательный| **Длинный**|Индекс нового объекта **гиперссылки** в коллекции **WebNavigationBarHyperlinks** .|

### <a name="return-value"></a>Возвращаемое значение

Hyperlink


## <a name="remarks"></a>Заметки

RelativePage может иметь одно из следующих констант [PbHlinkTargetType](pbhlinktargettype-enumeration-publisher.md) . Значение по умолчанию — **pbHlinkTargetTypeURL**.



| **pbHlinkTargetTypeEmail**|| **pbHlinkTargetTypeFirstPage**|| **pbHlinkTargetTypeLastPage**|| **pbHlinkTargetTypeNextPage**|| **pbHlinkTargetTypePageID**|| **pbHlinkTargetTypePreviousPage**|| **pbHlinkTargetTypeURL**|

