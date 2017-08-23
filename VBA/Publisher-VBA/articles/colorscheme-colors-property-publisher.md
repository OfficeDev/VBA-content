---
title: "Свойство ColorScheme.Colors (издатель)"
keywords: vbapb10.chm2686978
f1_keywords: vbapb10.chm2686978
ms.prod: publisher
api_name: Publisher.ColorScheme.Colors
ms.assetid: e6599096-3f99-e7ca-0c38-1cc7d4e0a1cd
ms.date: 06/08/2017
ms.openlocfilehash: 54729d81f55362016b5035cadd6766f6ed0547d5
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="colorschemecolors-property-publisher"></a>Свойство ColorScheme.Colors (издатель)

Возвращает объект **[ColorFormat](colorformat-object-publisher.md)** , представляющее цвет из указанного цветовая схема.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Цвета** ( **_ColorIndex (en)_**)

 переменная _expression_A, представляет собой объект- **ColorScheme** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|ColorIndex (en)|Обязательное свойство.| **PbSchemeColorIndex**| Цвет из схемы для возврата на основании его функции в схеме.|

### <a name="return-value"></a>Возвращаемое значение

ColorFormat


## <a name="remarks"></a>Заметки

Параметр ColorIndex может иметь одно из **[PbSchemeColorIndex](pbschemecolorindex-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

В следующем примере коллекции **ColorSchemes** и выполняет поиск цветовые схемы, где цвет папка сопоставляет цвета RGB значения 128.


```vb
Dim cscLoop As ColorScheme 
Dim colTemp As ColorFormat 
 
For Each cscLoop In Application.ColorSchemes 
 With cscLoop 
 Set colTemp = .Colors(ColorIndex:=pbSchemeColorFollowedHyperlink) 
 If colTemp.RGB = RGB(128, 0, 0) Then 
 Debug.Print "Color scheme '" &; .Name _ 
 &; "' has a followed hyperlink " _ 
 &; "color matching RGB(128, 0, 0)" 
 End If 
 End With 
Next cscLoop
```


