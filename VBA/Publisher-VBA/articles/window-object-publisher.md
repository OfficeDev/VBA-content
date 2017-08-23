---
title: "Объект Window (издатель)"
keywords: vbapb10.chm327679
f1_keywords: vbapb10.chm327679
ms.prod: publisher
api_name: Publisher.Window
ms.assetid: 342d77cd-5556-6ac3-a828-b1b60380f910
ms.date: 06/08/2017
ms.openlocfilehash: 7145579055403870923e6b7d49620b1b13869294
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="window-object-publisher"></a>Объект Window (издатель)

Представляет окно. Многие характеристики публикации, такие как полосы прокрутки и линейки, в действительности являются свойствами окна.
 


## <a name="example"></a>Пример

Свойство **[ActiveWindow](application-activewindow-property-publisher.md)** используется для возврата объекта **Window** . Следующий пример разворачивает активного окна.
 

 

```
Sub MaximizeWindow 
 ActiveWindow.WindowState = pbWindowStateMaximize 
End Sub
```

Свойство **[Caption](window-caption-property-publisher.md)** возвращает имена файлов и приложения активного окна. В следующем примере отображается сообщение с именем файла и имя приложения Microsoft Publisher.
 

 



```
Sub ShowFileApNames 
 MsgBox Windows(1).Caption 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Активация](window-activate-method-publisher.md)|
|[Перемещение](window-move-method-publisher.md)|
|[Чтобы изменить размер](window-resize-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](window-application-property-publisher.md)|
|[Заголовок](window-caption-property-publisher.md)|
|[Высота](window-height-property-publisher.md)|
|[HWND](window-hwnd-property-publisher.md)|
|[Слева](window-left-property-publisher.md)|
|[Родительский раздел](window-parent-property-publisher.md)|
|[Вверх](window-top-property-publisher.md)|
|[Visible](window-visible-property-publisher.md)|
|[Ширина](window-width-property-publisher.md)|
|[WindowState](window-windowstate-property-publisher.md)|

