---
title: "Объект буквицу (издатель)"
keywords: vbapb10.chm5570559
f1_keywords: vbapb10.chm5570559
ms.prod: publisher
api_name: Publisher.DropCap
ms.assetid: 7c6aeffe-cf25-a834-52de-5966df5e21d2
ms.date: 06/08/2017
ms.openlocfilehash: 4ef7547c5d5145cc49a9318b15f5630f35f00a1f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="dropcap-object-publisher"></a>Объект буквицу (издатель)

Представляет буквицы в начале абзаца.
 


## <a name="example"></a>Пример

Свойство **[буквицу](textrange-dropcap-property-publisher.md)** используется для возврата объекта **буквицу** . В следующем примере задается буквицы для первую букву каждого абзаца в первую фигуру на первой странице active публикации. Предполагается, что указанные форму текстовое поле и не другого типа фигуры.
 

 

```
Sub ApplyDropCap() 
 ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange _ 
 .DropCap.ApplyCustomDropCap Size:=3, Span:=3, Bold:=True 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[ApplyCustomDropCap](dropcap-applycustomdropcap-method-publisher.md)|
|[Очистить](dropcap-clear-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](dropcap-application-property-publisher.md)|
|[FontBold](dropcap-fontbold-property-publisher.md)|
|[FontColor](dropcap-fontcolor-property-publisher.md)|
|[FontItalic](dropcap-fontitalic-property-publisher.md)|
|[FontName](dropcap-fontname-property-publisher.md)|
|[LinesUp](dropcap-linesup-property-publisher.md)|
|[Родительский раздел](dropcap-parent-property-publisher.md)|
|[Размер](dropcap-size-property-publisher.md)|
|[Диапазон](dropcap-span-property-publisher.md)|

