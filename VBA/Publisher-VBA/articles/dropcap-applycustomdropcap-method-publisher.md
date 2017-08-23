---
title: "Метод DropCap.ApplyCustomDropCap (издатель)"
keywords: vbapb10.chm5505041
f1_keywords: vbapb10.chm5505041
ms.prod: publisher
api_name: Publisher.DropCap.ApplyCustomDropCap
ms.assetid: 906cf476-3826-8510-315f-425f6f50a92a
ms.date: 06/08/2017
ms.openlocfilehash: bbcad93b3d537d3e34c88331a9fad8b3a164283f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="dropcapapplycustomdropcap-method-publisher"></a>Метод DropCap.ApplyCustomDropCap (издатель)

Применяет пользовательского форматирования для первых знаков абзацев в рамке.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ApplyCustomDropCap** ( **_LinesUp_**, **_размер_**, **_диапазон_**, **_FontName_**, **_Полужирный_**, **_Курсив_**)

 переменная _expression_A, представляет собой объект- **буквицу** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|LinesUp|Необязательный| **Длинный**|Номер строки, чтобы переместить вверх буквицы. Значение по умолчанию равно 0. Максимальное количество не может быть больше, чем номер, введенный для аргумента размером менее одного.|
|Размер|Необязательный| **Длинный**|Размер символы буквицы в число строк. Значение по умолчанию равно 5.|
|Диапазон|Необязательный| **Длинный**|Число включенных в буквицы букв. Значение по умолчанию — 1.|
|FontName|Необязательный| **String**|Имя шрифта для форматирования буквицы. Значение по умолчанию — текущий шрифт.|
|Полужирный|Необязательный| **Boolean**| **Значение true,** полужирный шрифт буквицы. Значение по умолчанию — **False**.|
|Курсив|Необязательный| **Boolean**| **Значение true,** Чтобы применить курсивное буквицы. Значение по умолчанию — **False**.|

## <a name="example"></a>Пример

В этом примере форматов первые три буквы абзацы в указанном текстовом поле.


```vb
Sub CustDropCap() 
 
 ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange.DropCap _ 
 .ApplyCustomDropCap LinesUp:=1, Size:=6, Span:=3, _ 
 FontName:="Script MT Bold", Bold:=True, Italic:=True 
 
End Sub
```


