---
title: "Свойство Document.ColorScheme (издатель)"
keywords: vbapb10.chm196614
f1_keywords: vbapb10.chm196614
ms.prod: publisher
api_name: Publisher.Document.ColorScheme
ms.assetid: b7748b48-eff3-bdf0-e6ce-a9a2e788d0f7
ms.date: 06/08/2017
ms.openlocfilehash: 52d9fac69bdba47c3e6d48b347440dcd0b517e24
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentcolorscheme-property-publisher"></a>Свойство Document.ColorScheme (издатель)

Возвращает или задает объект **[ColorScheme](colorscheme-object-publisher.md)** , представляющий цвета схемы для указанной публикации. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ColorScheme**

 переменная _expression_A, представляющий объект **Document** .


### <a name="return-value"></a>Возвращаемое значение

ColorScheme


## <a name="example"></a>Пример

В этом примере отображается имя текущего цветовая схема active публикации.


```vb
With ActiveDocument.ColorScheme 
 MsgBox "The current color scheme is " &; .Name &; "." 
End With
```

В этом примере задается цветовая схема active публикации для «Alpine».




```vb
ActiveDocument.ColorScheme _ 
 = Application.ColorSchemes("Alpine")
```


