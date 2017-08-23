---
title: "Свойство Application.ColorSchemes (издатель)"
keywords: vbapb10.chm131080
f1_keywords: vbapb10.chm131080
ms.prod: publisher
api_name: Publisher.Application.ColorSchemes
ms.assetid: b991d8a2-d25d-839a-c14a-18cb6d126d33
ms.date: 06/08/2017
ms.openlocfilehash: d00046026eadb004b9687ab0814c7960323d8a2f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationcolorschemes-property-publisher"></a>Свойство Application.ColorSchemes (издатель)

Возвращает коллекцию **[ColorSchemes](colorschemes-object-publisher.md)** , представляющий доступные цветовые схемы.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ColorSchemes**

 переменная _expression_A, представляющий объект **приложения** .


### <a name="return-value"></a>Возвращаемое значение

ColorSchemes


## <a name="example"></a>Пример

В следующем примере коллекции **ColorSchemes** и отображает имя каждого цветовая схема и RGB значение цвета для последующей гиперссылки в каждой схемы.


```vb
Dim cscLoop As ColorScheme 
Dim cscAll As ColorSchemes 
 
Set cscAll = Application.ColorSchemes 
 
For Each cscLoop In cscAll 
 With cscLoop 
 Debug.Print "Color scheme: " &; .Name _ 
 &; " / Followed hyperlink color: " _ 
 &; .Colors(ColorIndex:=pbSchemeColorFollowedHyperlink).RGB 
 End With 
Next cscLoop
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

