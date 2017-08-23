---
title: "Свойство Font.StylisticSets (издатель)"
keywords: vbapb10.chm5374016
f1_keywords: vbapb10.chm5374016
ms.prod: publisher
api_name: Publisher.Font.StylisticSets
ms.assetid: 0d25fbf3-8d68-c10f-0d1b-526314700329
ms.date: 06/08/2017
ms.openlocfilehash: 7cad8f18fde0ba7470abcbf7285ebf16ab42e4fd
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fontstylisticsets-property-publisher"></a>Свойство Font.StylisticSets (издатель)

Возвращает или задает **Variant** , который представляет состояние свойства **StylisticSets** на символов в диапазон текста. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **StylisticSets**

 переменная _expression_A, представляющий объект **[Font](font-object-publisher.md)** .


## <a name="remarks"></a>Заметки

Свойство **StylisticSets** применяется от одного до 20 все более сложных наборов оформление стили для выбранного шрифта.

В следующей таблице показаны возможные значения для свойства **StylisticSets** и как они связаны с идентификаторами для Стилистические наборы пользовательского интерфейса (UI). Значение нуль (0) указывает, что применяется не стилистический набор.



|**Значение свойства StylisticSets**|**Стилистический набор идентификатор в пользовательском Интерфейсе**|
|:-----|:-----|
|0|0|
|1|1|
|2|2|
|4|3|
|8|4|
Номер стилистический наборов данных, доступных, может изменяться в зависимости от шрифта.


 **Примечание**  Свойство **StylisticSets** имеет значение только для шрифтов OpenType, которые содержат Стилистические наборы.


