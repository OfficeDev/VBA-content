---
title: "Свойство DropCap.FontColor (издатель)"
keywords: vbapb10.chm5505028
f1_keywords: vbapb10.chm5505028
ms.prod: publisher
api_name: Publisher.DropCap.FontColor
ms.assetid: 0c740ec7-05ac-b1fc-875c-cfd5a934c403
ms.date: 06/08/2017
ms.openlocfilehash: be7acec4bbbbaa8f0d83a369cefe6bd59c05741f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="dropcapfontcolor-property-publisher"></a>Свойство DropCap.FontColor (издатель)

Возвращает или задает объект **[ColorFormat](colorformat-object-publisher.md)** , представляющий цвет, применяемый к указанным буквицы.


## <a name="syntax"></a>Синтаксис

 _выражение_. **FontColor**

 переменная _expression_A, представляет собой объект- **буквицу** .


### <a name="return-value"></a>Возвращаемое значение

ColorFormat


## <a name="example"></a>Пример

В этом примере применяется цвета **[RGB](colorformat-rgb-property-publisher.md)** буквицы в элементе frame указанный текст. В этом примере предполагается, что указанный текст frame отформатирован буквицы.


```vb
Sub BoldDropCap() 
 With ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.DropCap 
 .FontBold = msoTrue 
 .FontColor.RGB = RGB(Red:=150, Green:=50, Blue:=180) 
 .FontItalic = msoTrue 
 .FontName = "Script MT Bold" 
 End With 
End Sub
```


