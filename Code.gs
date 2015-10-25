/*
 *  Script puts special signs, dashes and non-breaking spaces in the Russian text.
 * 
 *  Типограф расставляет спецзнаки, тире и неразрывные пробелы в русском тексте.
 *  При копипасте в гугл-документ неразрывные пробелы теряются.
 *  Скрипт заново меняет пробелы на неразрывные:
 *    после инициалов и чисел;
 *    после знаков номера, параграфа и копирайта;
 *    после "не" и "ни";
 *    после коротких предлогов и союзов;
 *    после коротких слов в начале предложения;
 *    после "гл.", "п.", "рис.", "ст.", "стр.", "таб." и "ч.";
 *    перед длинными тире, процентами, "б, "бы", "ж", "же", "ли", "ль";
 *    вокруг знаков равенства, сложения и умножения;
 *    внутри "и т.д.", "т.е." и т. п.
 *  Заменяет на длинное тире все черточки, сбоку от которых стоит пробел.
 *  Заменяет "+/-" на плюс-минус, "/=" на "не равно", " х " на умножение, точки на многоточие.
 *  Расставляет копирайты, торговые знаки и номера вместо их имитаций буквами.
 *  Убирает лишние пробелы.
 * 
 *  Ссылки по теме:
 *    http://artlebedev.ru/everything/izdal/spravochnik-izdatelya-i-avtora-2014/search/
 *    http://artlebedev.ru/kovodstvo/sections/62/
 *    https://en.wikipedia.org/wiki/Non-breaking_space
 *    https://ru.wikipedia.org/wiki/Неразрывный_пробел
 *    https://typograf.github.com
 *
 *  https://github.com/glebkema/google-docs-typograf
 */


function onOpen(e) {
  DocumentApp.getUi()
      .createAddonMenu()
      .addItem('Typograf', 'typograf')
      .addToUi();
}


function onInstall(e) {
  onOpen(e);
}


function clearSpaces(objBody) {
  var nbsp  = '\u00A0';
  var nbsp2 = '\\xA0';  // для поиска: \u00A0 не находит, \\u00A0 ругает
  var space = '\u0020';

  // лишние пробелы + смеси из обычных и неразрывных пробелов
  objBody.replaceText(space + space + '+',  space);
  objBody.replaceText(space + nbsp2,        nbsp);
  objBody.replaceText(nbsp2 + space,        nbsp);
  objBody.replaceText(nbsp2 + nbsp2 + '+',  nbsp);
}


function nbspAfter(objBody, strText) {
  var nbsp  = '\u00A0';
  var nbsp2 = '\\xA0';  // для поиска: \u00A0 не находит, \\u00A0 ругает
  var space = '\u0020';
  
  objBody.replaceText(space + strText + space,  space + strText + nbsp);
  objBody.replaceText(nbsp2 + strText + space,  nbsp  + strText + nbsp);
}


function nbspAfterComma(objBody, strText) {
  nbspAfter(objBody, strText);
  nbspAfter(objBody, strText + ',');
}


function nbspAfterDot(objBody, strText) {
  var nbsp  = '\u00A0';
  var nbsp2 = '\\xA0';  // для поиска: \u00A0 не находит, \\u00A0 ругает
  var space = '\u0020';
  
  var dot    = '\\.';

  // отодвигаем текст справа
  objBody.replaceText(space + strText + dot,  space + strText + '.' + nbsp);
  objBody.replaceText(nbsp2 + strText + dot,  nbsp  + strText + '.' + nbsp);
  objBody.replaceText(dot   + strText + dot,  '.'   + strText + '.' + nbsp);
  // !!!! упускаем случай, когда слева ничего нет (начало абзаца)
}


function nbspBefore(objBody, strText) {
  var nbsp  = '\u00A0';
  var nbsp2 = '\\xA0';  // для поиска: \u00A0 не находит, \\u00A0 ругает
  var space = '\u0020';

  objBody.replaceText(space + strText + space,  nbsp + strText + space);
  objBody.replaceText(space + strText + nbsp2,  nbsp + strText + nbsp );
}


function nbspBothSides(objBody, strSearch, strReplace) {
  var nbsp  = '\u00A0';
  var nbsp2 = '\\xA0';  // для поиска: \u00A0 не находит, \\u00A0 ругает
  var space = '\u0020';

  strReplace = nbsp + strReplace + nbsp;
  objBody.replaceText(space + strSearch + space,  strReplace);
  objBody.replaceText(space + strSearch + nbsp2,  strReplace);
  objBody.replaceText(nbsp2 + strSearch + space,  strReplace);
}


function nbspUnion(objBody, strText) {
  var dash   = '\u2014';
  var nbsp   = '\u00A0';
  var space  = '\u0020';
  
  objBody.replaceText(','  + space + strText + space,  ','  + space + strText + nbsp);
  
  // ???? почему то не работает :(
  // objBody.replaceText(dash + space + strText + space,  dash + space + strText + nbsp);
}


function typograf() {
  var body = DocumentApp.getActiveDocument().getBody();
  
  /****
  !!!! отодвигаем текст сбоку, даже если там есть пробел + лишнее убираем в конце
  !!!! точку надо экранировать
  !!!! нужны две косые черты, поскольку одна пропадает при обработке выражения
  !!!! не принимает выражения '  ' и ' т(?=\\.)'
  !!!! при поиске '\xA0' находит, а '\u00A0' ругает
  !!!! для объединения строк используем + вместо &

  xA0 u00A0 no-break space
  xA7 u00A7 section
  xD7 u00D7 multiplication
  xA9 u00A9 copyright
  xAE u00AE registered
  xB0 u00B0 degree
  xB1 u00B1 plus-minus
      u2026 horizontal ellipsis (многоточие)
      u2116 numero
      u2117 sound recording copyright
      u2122 trademark
      u2260 not equal to

  x2D u002D hyphen-minus
      u2010 hyphen 
      u2011 non-breaking hyphen
      u2012 figure dash
      u2013 en dash
      u2014 em dash
      u2015 horizontal bar
  ****/

  var dash   = '\u2014';
  var dots   = '\u2026';
  var nbsp   = '\u00A0';
  var nbsp2  =  '\\xA0';  // для поиска: \u00A0 не находит, \\u00A0 ругает
  var mult   = '\u00D7';
  var numero = '\u2116';
  var space  = '\u0020';
  var trade  = '\u2122';
  
  var dot    = '\\.';
  var left   = '\\(';
  var plus   = '\\+';
  var right  = '\\)';

  var dashWrong = '\u002D\u2010\u2011\u2012\u2013\u2015';
  var dashAll   = '[' + dashWrong + dash + ']';
  dashWrong     = '[' + dashWrong + ']';

  clearSpaces(body); 
  
  // пробел перед точкой, запятой и точкой с запятой
  body.replaceText(space + dot,  '.');
  body.replaceText(space + ',',  ',');
  body.replaceText(space + ';',  ';');

  // многоточие
  body.replaceText(dot  + dot  + '+',  dots);  // две и более точек
  body.replaceText(dots + dot  + '+',  dots);  // многоточие и точки
  body.replaceText(dots + dots + '+',  dots);  // несколько многоточий
  body.replaceText(dot  + dots,        dots);  // точка и многоточие (2 точки уже обработаны)
  
  // копирайт + из русских букв + отодвигаем текст справа
  body.replaceText(left + '[cCсС]' + right, '\xA9' + nbsp); 

  // sound recording copyright / published / phonorecord sign + из русских букв + отодвигаем
  body.replaceText(left + '[pPрР]' + right, '\u2117' + nbsp); 

  // зарегистрированный товарный знак
  body.replaceText(left + '[rR]' + right, '\xAE'); 

  // товарный знак + из русских букв
  body.replaceText(left + 'tm' + right, trade); 
  body.replaceText(left + 'TM' + right, trade); 
  body.replaceText(left + 'тм' + right, trade); 
  body.replaceText(left + 'ТМ' + right, trade); 

  /**** математика 
  // NB: если пробелы стоят с обеих сторон
  ****/

  // умножение
  nbspBothSides(body, mult, mult);
  nbspBothSides(body,  'x', mult);
  nbspBothSides(body,  'X', mult);
  nbspBothSides(body,  'х', mult);
  nbspBothSides(body,  'Х', mult);
  
  // плюс и равно
  nbspBothSides(body, plus, '+');
  nbspBothSides(body,  '=', '=');
  
  // плюс-минус + неразрывный справа
  body.replaceText(plus       + dashAll, '\xB1' + nbsp); 
  body.replaceText(plus + '/' + dashAll, '\xB1' + nbsp); 
  
  // не равно
  body.replaceText('/=|=/', '\u2260'); 
  
  // тире -> слева неразрывный пробел + справа обычный
  // за тире принимаем любой дефис или тире, у которого хотя бы с 1 бока стоит пробел
  // могут быть: разные дефисы и тире
  //              разные пробелы по бокам: _-_, _-n, n-_, n-n
  //              пробел только с 1 бока:  a-_, a-n, _-a, n-a
  //              несколько дефисов и тире подряд и вперемешку
  // несколько дефисов подряд
  body.replaceText('--+',              nbsp + dash);
  // один вид тире + если слева пробел то неразрывный
  body.replaceText(space + dashAll,    nbsp + dash);
  body.replaceText(nbsp2 + dashWrong,  nbsp + dash);
  // убираем дефисы и тире справа
  body.replaceText(dash  + dashAll + '+',     dash);
  // один вид тире + если справа пробел то обычный
  body.replaceText(dashWrong + space,  dash + space);
  body.replaceText(dashAll   + nbsp2,  dash + space);
  // повторяющиеся тире
  body.replaceText(dash + nbsp2 + dash,   nbsp + dash + space);
  body.replaceText(dash + space + dash,   nbsp + dash + space);
  // пробелы с обеих сторон + слева 1-2 неразрывных + справа 1-2 обычных
  body.replaceText(nbsp2 + dash,          nbsp + dash + space);
  body.replaceText(        dash + space,  nbsp + dash + space);
  
  /**** нужно предлагать выбор кавычек + разобраться со вложенными кавычками
  http://www.artlebedev.ru/kovodstvo/sections/62/
  https://ru.wikipedia.org/wiki/%D0%9A%D0%B0%D0%B2%D1%8B%D1%87%D0%BA%D0%B8
  // парные кавычки
  body.replaceText('“', '«'); 
  body.replaceText('”', '»'); 
  // непарные кавычки надо опознать
  body.replaceText(' "', ' «'); 
  // то же для начала строки
  body.replaceText('" ', '» '); 
  // то же для конца строки
  body.replaceText('"\\.', '».'); 
  body.replaceText('",', '»,'); 
  ****/
  
  // номер
  body.replaceText(numero,                        numero + nbsp);  // отодвигаем текст справа
  body.replaceText(space + 'No' + space,  space + numero + nbsp); 
  body.replaceText(nbsp2 + 'No' + space,  nbsp  + numero + nbsp); 
  body.replaceText(space + 'N'  + space,  space + numero + nbsp); 
  body.replaceText(nbsp2 + 'N'  + space,  nbsp  + numero + nbsp); 

  // после параграфа + отодвигаем текст справа
  body.replaceText('\\xA7', '\xA7' + nbsp); 
  
  // после инициалов
  for (var i=1040; i<1072; i++) {
    if (i < 1066 | i > 1068) { 
      nbspAfterDot(body, String.fromCharCode(i));
      // проверка: body.appendParagraph(String.fromCharCode(i)); 
    }
  }

  // после чисел
  for (var i = 0; i < 10; i++) {
    body.replaceText(i + space, i + nbsp); 
  }
  
  // проценты + отодвигаем текст слева
  body.replaceText('%', nbsp + '%'); 
  // проверяем двойной знак процентов поле чистки лишних пробелов
  
  // если после обработки чисел перед г. остался пробел, то это не год или граммы, а город
  body.replaceText(' г\\. ', ' г.' + nbsp); 
  
  // т., и т. д.; и т. п.; т. е.; т. к.; т. н.; т. о.
  body.replaceText('и т\\.', 'и' + nbsp + 'т.');
  nbspAfterDot(body, 'т');

  // церковь: https://ru.wikipedia.org/wiki/Церковные_сокращения 
  nbspAfterDot(body, 'кн');  nbspAfterDot(body, 'Кн');
  nbspAfterDot(body, 'св');  nbspAfterDot(body, 'Св');
  
  // адрес: ул., д., кв., к., т., м., ф.
  
  // глава, пункт, рисунок, статья, страница, таблица, часть
  // книга = князь
  nbspAfterDot(body, 'гл');
  nbspAfterDot(body, 'п');
  nbspAfterDot(body, 'рис');
  nbspAfterDot(body, 'ст');
  nbspAfterDot(body, 'стр');
  nbspAfterDot(body, 'табл');
  nbspAfterDot(body, 'ч');

  // слово из 1 буквы в начале предложения
  nbspAfter(body, 'В');
  nbspAfter(body, 'К');
  nbspAfter(body, 'С');
  nbspAfter(body, 'У');

  nbspAfterComma(body, 'А');
  nbspAfterComma(body, 'И');
  nbspAfterComma(body, 'О');
  nbspAfterComma(body, 'Я');

  // слово из 2 букв в начале предложения
  nbspAfter(body, 'Ан');
  nbspAfter(body, 'Во');
  nbspAfter(body, 'За');
  nbspAfter(body, 'Из');
  nbspAfter(body, 'Ко');
  nbspAfter(body, 'На');
  nbspAfter(body, 'Не');
  nbspAfter(body, 'Ни');
  nbspAfter(body, 'Ор');
  nbspAfter(body, 'От');
  nbspAfter(body, 'По');
  nbspAfter(body, 'Со');
  nbspAfter(body, 'Уж');

  nbspAfterComma(body, 'Ах');
  nbspAfterComma(body, 'Вы');
  nbspAfterComma(body, 'Да');
  nbspAfterComma(body, 'До');
  nbspAfterComma(body, 'Ее');
  nbspAfterComma(body, 'Её');
  nbspAfterComma(body, 'Ею');
  nbspAfterComma(body, 'Им');
  nbspAfterComma(body, 'Их');
  nbspAfterComma(body, 'Мы');
  nbspAfterComma(body, 'Но');
  nbspAfterComma(body, 'Ну');
  nbspAfterComma(body, 'Он');
  nbspAfterComma(body, 'Ох');
  nbspAfterComma(body, 'Та');
  nbspAfterComma(body, 'Те');
  nbspAfterComma(body, 'То');
  nbspAfterComma(body, 'Ту');
  nbspAfterComma(body, 'Ты');
  nbspAfterComma(body, 'Эх');

  /* надо ???
  nbspAfterComma(body, 'Ас');
  nbspAfterComma(body, 'Юг');
  */

  // слово из 3 букв в начале предложения
  nbspAfter(body, 'Вон');
  nbspAfter(body, 'Для');
  nbspAfter(body, 'Изо');
  nbspAfter(body, 'Над');
  nbspAfter(body, 'Ото');
  nbspAfter(body, 'Под');
  nbspAfter(body, 'При');
  nbspAfter(body, 'Про');
  nbspAfter(body, 'Чей');
  nbspAfter(body, 'Чем');
  nbspAfter(body, 'Что');
  nbspAfter(body, 'Чью');
  nbspAfter(body, 'Чья');
  
  nbspAfterComma(body, 'Вам');
  nbspAfterComma(body, 'Вас');
  nbspAfterComma(body, 'Вот');
  nbspAfterComma(body, 'Все');
  nbspAfterComma(body, 'Всю');
  nbspAfterComma(body, 'Вся');
  nbspAfterComma(body, 'Ему');
  nbspAfterComma(body, 'Или');
  nbspAfterComma(body, 'Как');
  nbspAfterComma(body, 'Кем');
  nbspAfterComma(body, 'Мой');
  nbspAfterComma(body, 'Нам');
  nbspAfterComma(body, 'Нас');
  nbspAfterComma(body, 'Наш');
  nbspAfterComma(body, 'Она');
  nbspAfterComma(body, 'Они');
  nbspAfterComma(body, 'Оно');
  nbspAfterComma(body, 'Рад');
  nbspAfterComma(body, 'Сей');
  nbspAfterComma(body, 'Сию');
  nbspAfterComma(body, 'Сия');
  nbspAfterComma(body, 'Так');
  nbspAfterComma(body, 'Там');
  nbspAfterComma(body, 'Тем');
  nbspAfterComma(body, 'Тех');
  nbspAfterComma(body, 'Той');
  nbspAfterComma(body, 'Тот');
  nbspAfterComma(body, 'Эта');
  nbspAfterComma(body, 'Эти');
  nbspAfterComma(body, 'Это');
  nbspAfterComma(body, 'Эту');

  // предлоги и союзы из 1 буквы
  nbspAfter(body, 'а');
  nbspAfter(body, 'в');
  nbspAfter(body, 'к');
  nbspAfter(body, 'о');
  nbspAfter(body, 'с');
  nbspAfter(body, 'у');

  // предлоги и частицы из 2 букв
  nbspAfter(body, 'во');
  nbspAfter(body, 'до');
  nbspAfter(body, 'за');
  nbspAfter(body, 'из');
  nbspAfter(body, 'ко');
  nbspAfter(body, 'на');
  nbspAfter(body, 'не');
  nbspAfter(body, 'ни');
  nbspAfter(body, 'но');
  nbspAfter(body, 'ну');
  nbspAfter(body, 'от');
  nbspAfter(body, 'по');
  nbspAfter(body, 'со');
  nbspAfter(body, 'то');

  // слово после запятой
  nbspUnion(body, 'вам');
  nbspUnion(body, 'вас');
  nbspUnion(body, 'ваш');
  nbspUnion(body, 'вы');
  nbspUnion(body, 'да');
  nbspUnion(body, 'его');
  nbspUnion(body, 'ее');
  nbspUnion(body, 'её');
  nbspUnion(body, 'если');
  nbspUnion(body, 'и');
  nbspUnion(body, 'или');
  nbspUnion(body, 'им');
  nbspUnion(body, 'их');
  nbspUnion(body, 'как');
  nbspUnion(body, 'кем');
  nbspUnion(body, 'мы');
  nbspUnion(body, 'нам');
  nbspUnion(body, 'нас');
  nbspUnion(body, 'он');
  nbspUnion(body, 'она');
  nbspUnion(body, 'они');
  nbspUnion(body, 'оно');
  nbspUnion(body, 'та');
  nbspUnion(body, 'так');
  nbspUnion(body, 'там');
  nbspUnion(body, 'те');
  nbspUnion(body, 'тем');
  nbspUnion(body, 'тех');
  nbspUnion(body, 'то');
  nbspUnion(body, 'той');
  nbspUnion(body, 'тот');
  nbspUnion(body, 'чей');
  nbspUnion(body, 'чем');
  nbspUnion(body, 'что');
  nbspUnion(body, 'чью');
  nbspUnion(body, 'эта');
  nbspUnion(body, 'эти');
  nbspUnion(body, 'это');

  // перед бы, же, ли
  nbspBefore(body, 'б');
  nbspBefore(body, 'бы');
  nbspBefore(body, 'ж');
  nbspBefore(body, 'же');
  nbspBefore(body, 'ли');
  nbspBefore(body, 'ль');

  clearSpaces(body); 
  
  // проверяем двойной знак процентов поле чистки лишних пробелов
  body.replaceText('%' + nbsp2 + '%', '%%'); 

  // пробелы в начале и в конце строки
 
  // пустые строки
}
