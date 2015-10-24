### Скрипт расставляет спецзнаки, тире и неразрывные пробелы в русскоязычных гугл-документах.

При копипасте в гугл-документ неразрывные пробелы теряются. Скрипт заново меняет пробелы на неразрывные:

* после инициалов и чисел;
* после знаков номера, параграфа и копирайта;
* после "не" и "ни";
* после коротких предлогов и союзов;
* после коротких слов в начале предложения;
* после "гл.", "п.", "рис.", "ст.", "стр.", "таб." и "ч.";
* перед длинными тире, процентами, "б, "бы", "ж", "же", "ли", "ль";
* вокруг знаков равенства, сложения и умножения;
* внутри "и т.д.", "т.е." и т. п.

Заменяет:

* дефисы и тире, сбоку от которых стоит пробел, на длинное тире;
* цепочки точек на многоточия;
* "+/-" и "+-" на плюс-минус;
* "/=" и "=/" на "не равно";
* " х " на знак умножения;
* "(с)", "(р)" и "(тм)", набранные русскими и латинскими буквами, на соответствующие знаки;
* " N " и " No " на знак номера.

Убирает лишние пробелы.
