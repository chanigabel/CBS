# טבלת בדיקת דרישות - עברית

מבוסס על הגיליון `וולידציות לשדות בקובץ` מתוך `BAKAROT.xlsx`.

## מקרא
- ירוק: דרישת המסמך / דרישות בסיס
- כחול: ממומש
- אדום: חסר
- צהוב: מעבר לדרישה
- אפור: הערות

## סיכום מנהלים
- המערכת הפעילה היא נתיב ה־Web/API Dataset Pipeline.
- נתיבי Excel ישירים קיימים רק לצורך תאימות היסטורית, ואינם נתיב הריצה הפעיל.
- הפערים המשמעותיים ביותר הם אימות מילון סוג מוסד עבור SugMosad, בדיקת מרשם אוכלוסין ומוסדות קשורים עבור MisparZehut, כלל גיל כניסה מינימלי לפי סוג מוסד, ואי־התאמת הסקופ של YomKnisa.
- ההתנהגויות הנוספות הבולטות הן ניקוי מזהים, ריפוד ספרות לתעודת זהות, העברת מזהה לא תקין לדרכון, תיקון תאריכים, זיהוי כפילויות, סיכום קומפקטי של דוחות עיבוד, והקצאת SugMosad לפי טווחי פעולה.

## בדיקת שדות
### MosadID
- גיליונות: AnasheyTzevet
DayarimYahidim
MeshkeyBayt
- נדרש לפי המסמך: חובה
- דרישות המסמך: חובה בייצוא בלבד
- ממומש: כן / חלקי / מעבר לדרישה
- מצב במערכת: קיים ייצוא של השדה בכל שלוש הלשוניות; המערכת מזריקה את מזהה המוסד מה-session לשורות התצוגה והייצוא; דוח העיבוד מסכם שדות ייצוא חסרים.
- חסר: לא נמצא
- מעבר לדרישה: אין חסימה קשיחה אלא דיווח ואזהרות בלבד.
- נתיב פעיל או legacy: פעיל ב־Web/API Dataset Pipeline; קיימת גם התאמה ישנה היכן שרלוונטי.
- קוד / בדיקות: פעיל ב־Web/API Dataset Pipeline. בדיקות: tests/test_institution_report_validator.py, tests/webapp/test_api_institution.py, tests/webapp/test_api_process_file.py. קיימת גם התאמה ישנה ב־export_engine.
- הערות סיכון: פעיל ב־Web/API Dataset Pipeline. בדיקות: tests/test_institution_report_validator.py, tests/webapp/test_api_institution.py, tests/webapp/test_api_process_file.py. קיימת גם התאמה ישנה ב־export_engine.

### SugMosad
- גיליונות: AnasheyTzevet
DayarimYahidim
MeshkeyBayt
- נדרש לפי המסמך: חובה
- דרישות המסמך: ערכים נומריים
אורך מזהה מוסד לפחות 3 תווים
לוודא שהמוסד קיים במילון סוג מוסד 
- ממומש: כן / חלקי / מעבר לדרישה
- מצב במערכת: קיים ייצוא של השדה בכל שלוש הלשוניות; המערכת מזריקה את סוג המוסד הפעיל לשורות; קיימת בדיקה לנומריות ולמינימום של 3 תווים; דוח העיבוד מסכם ערכי ייצוא חסרים.
- חסר: לא נמצאה בדיקת חברות במילון סוג מוסד בקוד הפעיל; ב-validator קיים סימון TODO מפורש.
- מעבר לדרישה: קיימת תמיכה בהקצאה לפי קובץ, לפי גיליון ולפי שורות מסומנות, וכן ערך ברירת־מחדל מה-session.
- נתיב פעיל או legacy: פעיל ב־Web/API Dataset Pipeline; קיימת גם התאמה ישנה היכן שרלוונטי.
- קוד / בדיקות: פעיל ב־Web/API Dataset Pipeline. נבדקו: institution_report_validator.py, standardization_service.py, export_service.py, report_collectors.py. לא נמצאה בדיקת מילון.
- הערות סיכון: פעיל ב־Web/API Dataset Pipeline. נבדקו: institution_report_validator.py, standardization_service.py, export_service.py, report_collectors.py. לא נמצאה בדיקת מילון.

### MisparDiraBeMosad
- גיליונות: AnasheyTzevet
MeshkeyBayt
- נדרש לפי המסמך: חובה
- דרישות המסמך: ערכים נומריים
- ממומש: כן / חלקי / מעבר לדרישה
- מצב במערכת: קיים ייצוא של השדה ב־AnasheyTzevet וב־MeshkeyBayt; ה-validator מקבל ערך נומרי כאשר הוא קיים ומדלג על DayarimYahidim.
- חסר: לפי המסמך השדה נדרש ביצוא, אך ה-validator הפעיל מתייחס אליו כאופציונלי; ערכים חסרים נחשפים רק בעקיפין דרך הייצוא והדוח.
- מעבר לדרישה: אין ולידציה נוספת מעבר לבדיקה שהערך נומרי.
- נתיב פעיל או legacy: פעיל ב־Web/API Dataset Pipeline; קיימת גם התאמה ישנה היכן שרלוונטי.
- קוד / בדיקות: פעיל ב־Web/API Dataset Pipeline. קיימת אי־התאמה בין דרישת המסמך לבין רמת האכיפה ב-validator.
- הערות סיכון: פעיל ב־Web/API Dataset Pipeline. קיימת אי־התאמה בין דרישת המסמך לבין רמת האכיפה ב-validator.

### ShemPrati
- גיליונות: AnasheyTzevet
DayarimYahidim
MeshkeyBayt
- נדרש לפי המסמך: None
- דרישות המסמך: ?
- ממומש: כן / חלקי / מעבר לדרישה
- מצב במערכת: ניקוי השמות יוצר first_name_corrected; ה-export schema כולל את השדה בכל הלשוניות; ה-validator הפעיל מסמן שם פרטי חסר.
- חסר: לפי המסמך השדה אינו חובה / לא הוגדר, אך המערכת מתייחסת אליו כחובה ומדווחת שגיאה כאשר הוא ריק.
- מעבר לדרישה: הודעת המצב ברמת השורה מופיעה גם ב־_validation_status ב-UI.
- נתיב פעיל או legacy: פעיל ב־Web/API Dataset Pipeline; קיימת גם התאמה ישנה היכן שרלוונטי.
- קוד / בדיקות: פעיל ב־Web/API Dataset Pipeline. דוח העיבוד נשאר קומפקטי ואינו משקף כל הודעת validator ברמת שורה.
- הערות סיכון: פעיל ב־Web/API Dataset Pipeline. דוח העיבוד נשאר קומפקטי ואינו משקף כל הודעת validator ברמת שורה.

### ShemMishpaha
- גיליונות: AnasheyTzevet
DayarimYahidim
MeshkeyBayt
- נדרש לפי המסמך: None
- דרישות המסמך: ?
- ממומש: כן / חלקי / מעבר לדרישה
- מצב במערכת: ניקוי השמות יוצר last_name_corrected; ה-export schema כולל את השדה בכל הלשוניות; ה-validator הפעיל מסמן שם משפחה חסר.
- חסר: לפי המסמך השדה אינו חובה / לא הוגדר, אך המערכת מתייחסת אליו כחובה ומדווחת שגיאה כאשר הוא ריק.
- מעבר לדרישה: הודעת המצב ברמת השורה מופיעה גם ב־_validation_status ב-UI.
- נתיב פעיל או legacy: פעיל ב־Web/API Dataset Pipeline; קיימת גם התאמה ישנה היכן שרלוונטי.
- קוד / בדיקות: פעיל ב־Web/API Dataset Pipeline. דוח העיבוד נשאר קומפקטי ואינו משקף כל הודעת validator ברמת שורה.
- הערות סיכון: פעיל ב־Web/API Dataset Pipeline. דוח העיבוד נשאר קומפקטי ואינו משקף כל הודעת validator ברמת שורה.

### ShemHaAv
- גיליונות: AnasheyTzevet
DayarimYahidim
MeshkeyBayt
- נדרש לפי המסמך: None
- דרישות המסמך: ?
- ממומש: כן / חלקי / מעבר לדרישה
- מצב במערכת: ניקוי השמות יוצר father_name_corrected; ה-export schema כולל את השדה בכל הלשוניות.
- חסר: לא נמצאה ולידציה ייעודית לשם האב בקוד הפעיל.
- מעבר לדרישה: השדה נכלל בייצוא גם בלי כלל בדיקה מפורש במסמך.
- נתיב פעיל או legacy: פעיל ב־Web/API Dataset Pipeline; קיימת גם התאמה ישנה היכן שרלוונטי.
- קוד / בדיקות: חיפשתי ב־institution_report_validator.py ולא נמצא _validate_shem_haav.
- הערות סיכון: חיפשתי ב־institution_report_validator.py ולא נמצא _validate_shem_haav.

### MisparZehut
- גיליונות: AnasheyTzevet
DayarimYahidim
MeshkeyBayt
- נדרש לפי המסמך: חובה
- דרישות המסמך: תקינות תעודת זהות 
תקינות ספרת ביקורת
הצלבה מול מרשם אוכלוסין 
לוודא שת"ז יוניקית בכל הקובץ שהתקבל ושת"ז לא מופיעה יותר מפעם אחת בכל אחד מהגיליונות
לוודא שת"ז אינה מדווחת באחד מהמוסדות הקשורים למוסד המדווח 
- ממומש: כן / חלקי / מעבר לדרישה
- מצב במערכת: IdentifierEngine מסיר מקפים, מרחיב ל־9 ספרות, בודק ספרת ביקורת, דוחה מזהים עם אפסים בלבד, ויכול להעביר מזהה לא תקין לדרכון; ה-validator בודק חסר וכפילויות בתוך גיליון ובין גיליונות; דוח העיבוד מסכם מזהים חסרים ומזהים לא תקינים.
- חסר: לא נמצאה הצלבה מול מרשם אוכלוסין ולא נמצאה בדיקה מול מוסדות קשורים; בשני המקרים יש סימון TODO מפורש ב-validator.
- מעבר לדרישה: יש תמיכה ב־9999 כסימון למזהה חסר, בהעברה לדרכון כאשר הת״ז לא תקינה, ובהצגת כפילויות ב־_validation_status.
- נתיב פעיל או legacy: פעיל ב־Web/API Dataset Pipeline; קיימת גם התאמה ישנה היכן שרלוונטי.
- קוד / בדיקות: פעיל ב־Web/API Dataset Pipeline. בדיקות: tests/test_identifier_engine.py, tests/test_institution_report_validator.py, tests/webapp/test_api_process_file.py, tests/webapp/test_api_institution.py. קיימת גם התאמת ייצוא ישנה.
- הערות סיכון: פעיל ב־Web/API Dataset Pipeline. בדיקות: tests/test_identifier_engine.py, tests/test_institution_report_validator.py, tests/webapp/test_api_process_file.py, tests/webapp/test_api_institution.py. קיימת גם התאמת ייצוא ישנה.

### Darkon
- גיליונות: AnasheyTzevet
DayarimYahidim
MeshkeyBayt
- נדרש לפי המסמך: None
- דרישות המסמך: None
- ממומש: כן / חלקי / מעבר לדרישה
- מצב במערכת: passport_corrected נוצר ונוקה ע״י IdentifierEngine; ה-export schema כולל את Darkon בכל הלשוניות.
- חסר: לא נמצאה ולידציה ייעודית לדרכון במסלול הפעיל.
- מעבר לדרישה: דרכון משמש כיעד חלופי כאשר ת״ז חסרה או לא תקינה.
- נתיב פעיל או legacy: פעיל ב־Web/API Dataset Pipeline; קיימת גם התאמה ישנה היכן שרלוונטי.
- קוד / בדיקות: פעיל ב־Web/API Dataset Pipeline. אין התנהגות legacy-only נפרדת לשדה זה מעבר למנוע הייצוא הישן המקביל.
- הערות סיכון: פעיל ב־Web/API Dataset Pipeline. אין התנהגות legacy-only נפרדת לשדה זה מעבר למנוע הייצוא הישן המקביל.

### Min
- גיליונות: AnasheyTzevet
DayarimYahidim
MeshkeyBayt
- נדרש לפי המסמך: None
- דרישות המסמך: קוד מין 
1- זכר 
2- נקבה 
- ממומש: כן / חלקי / מעבר לדרישה
- מצב במערכת: GenderEngine ממפה ערכי מין טקסטואליים ל־1/2; ה-validator מקבל רק 1 או 2 כאשר ערך קיים; ה-export schema כולל את Min בכל הלשוניות.
- חסר: לא נמצא פער משמעותי ביחס למסמך.
- מעבר לדרישה: המערכת ממפה מגוון רחב של כינויים טקסטואליים ל־1/2 ומחזירה הודעות שורה עבור ערכים לא תקינים.
- נתיב פעיל או legacy: פעיל ב־Web/API Dataset Pipeline; קיימת גם התאמה ישנה היכן שרלוונטי.
- קוד / בדיקות: פעיל ב־Web/API Dataset Pipeline. בדיקות: tests/test_gender_engine.py, tests/test_institution_report_validator.py.
- הערות סיכון: פעיל ב־Web/API Dataset Pipeline. בדיקות: tests/test_gender_engine.py, tests/test_institution_report_validator.py.

### ShnatLida
- גיליונות: AnasheyTzevet
DayarimYahidim
MeshkeyBayt
- נדרש לפי המסמך: None
- דרישות המסמך: תקינות פורמט שנת לידה
שנת לידה מקסימלי 1906
שנת לידה אינה עתידית
- ממומש: כן / חלקי / מעבר לדרישה
- מצב במערכת: DateEngine מפענח תאריכי לידה מפוצלים ומשולבים, מרחיב שנה דו־ספרתית, בודק תאריך אמיתי, אוכף שנת לידה מינימלית 1906, ומעדכן birth_date_status; ה-validator בודק גם רכיבים נדרשים.
- חסר: לפי המסמך השדה אינו חובה, אך ה-validator מתייחס להיעדר שנה / חודש / יום כשגיאה.
- מעבר לדרישה: המנוע מוסיף גם אזהרת גיל מעל 100, וברשת ה־UI מוצגות הודעות _validation_status ברמת שורה.
- נתיב פעיל או legacy: פעיל ב־Web/API Dataset Pipeline; קיימת גם התאמה ישנה היכן שרלוונטי.
- קוד / בדיקות: פעיל ב־Web/API Dataset Pipeline. בדיקות: tests/test_date_engine.py, tests/test_institution_report_validator.py, tests/test_plain_date_columns.py, tests/test_per_field_date_detection.py.
- הערות סיכון: פעיל ב־Web/API Dataset Pipeline. בדיקות: tests/test_date_engine.py, tests/test_institution_report_validator.py, tests/test_plain_date_columns.py, tests/test_per_field_date_detection.py.

### HodeshLida
- גיליונות: AnasheyTzevet
DayarimYahidim
MeshkeyBayt
- נדרש לפי המסמך: None
- דרישות המסמך: חודש לידה נומרי 
טווח ערכים תקינים בין 1 - 12
- ממומש: כן / חלקי / מעבר לדרישה
- מצב במערכת: DateEngine בודק שחודש לידה הוא מספרי ובטווח 1-12; ערכים לא תקינים נרשמים כ-corrected ריק; ה-validator בודק חסר / לא־נומרי / חריגה מהטווח.
- חסר: לפי המסמך השדה אינו חובה, אך ה-validator מתייחס להיעדר חודש כשגיאה.
- מעבר לדרישה: מנוע התאריכים המלא מחיל גם כללי תאריך מדויק וכללי עתידיות.
- נתיב פעיל או legacy: פעיל ב־Web/API Dataset Pipeline; קיימת גם התאמה ישנה היכן שרלוונטי.
- קוד / בדיקות: פעיל ב־Web/API Dataset Pipeline. בדיקות: tests/test_date_engine.py, tests/test_institution_report_validator.py.
- הערות סיכון: פעיל ב־Web/API Dataset Pipeline. בדיקות: tests/test_date_engine.py, tests/test_institution_report_validator.py.

### YomLida
- גיליונות: AnasheyTzevet
DayarimYahidim
MeshkeyBayt
- נדרש לפי המסמך: None
- דרישות המסמך: יום לידה נומרי 
טווח ערכים תקינים בין 1 - 31 תלוי חודש לידה
- ממומש: כן / חלקי / מעבר לדרישה
- מצב במערכת: DateEngine בודק שיום לידה הוא מספרי ובטווח 1-31; ערכים לא תקינים נרשמים כ-corrected ריק; ה-validator בודק חסר / לא־נומרי / חריגה מהטווח.
- חסר: לפי המסמך השדה אינו חובה, אך ה-validator מתייחס להיעדר יום כשגיאה.
- מעבר לדרישה: מנוע התאריכים המלא מחיל גם כללי תאריך מדויק וכללי עתידיות.
- נתיב פעיל או legacy: פעיל ב־Web/API Dataset Pipeline; קיימת גם התאמה ישנה היכן שרלוונטי.
- קוד / בדיקות: פעיל ב־Web/API Dataset Pipeline. בדיקות: tests/test_date_engine.py, tests/test_institution_report_validator.py.
- הערות סיכון: פעיל ב־Web/API Dataset Pipeline. בדיקות: tests/test_date_engine.py, tests/test_institution_report_validator.py.

### shnatknisa
- גיליונות: AnasheyTzevet
DayarimYahidim
MeshkeyBayt
- נדרש לפי המסמך: None
- דרישות המסמך: שנת כניסה ערכים נומריים 
שנת כניסה לא מאוחר מ 31.12 לשנת פקידה 
שנת כניסה אינה קודמת לתאריך הלידה 
שנת הכניסה אינה קודמת לגיל כניסה מינימלי המוגדר לפי סוג מוסד 
- ממומש: כן / חלקי / מעבר לדרישה
- מצב במערכת: DateEngine בודק שנת כניסה מול שנת המפקד ומול תאריך הלידה; standardization מוסיף אזהרה כאשר הכניסה קודמת ללידה; ה-validator מטפל ב־shnatknisa כחלק מבדיקת תאריך הכניסה.
- חסר: לפי המסמך נדרש גיל כניסה מינימלי לפי סוג מוסד, אך לא נמצא מיפוי פעיל של SugMosad → גיל מינימום.
- מעבר לדרישה: רכיבי תאריך מפוצל נשמרים, והמערכת מוסיפה טקסט אזהרה ל־entry_date_status.
- נתיב פעיל או legacy: פעיל ב־Web/API Dataset Pipeline; קיימת גם התאמה ישנה היכן שרלוונטי.
- קוד / בדיקות: פעיל ב־Web/API Dataset Pipeline. ב-validator יש סימון TODO מפורש לכלל גיל הכניסה. בדיקות: tests/test_date_engine.py, tests/test_institution_report_validator.py.
- הערות סיכון: פעיל ב־Web/API Dataset Pipeline. ב-validator יש סימון TODO מפורש לכלל גיל הכניסה. בדיקות: tests/test_date_engine.py, tests/test_institution_report_validator.py.

### Hodeshknisa
- גיליונות: AnasheyTzevet
DayarimYahidim
MeshkeyBayt
- נדרש לפי המסמך: None
- דרישות המסמך: חודש כניסה ערכים נומריים 
בין 1  לחודש 12 לשנת מפקד 
- ממומש: כן / חלקי / מעבר לדרישה
- מצב במערכת: DateEngine בודק שחודש כניסה הוא מספרי ובטווח 1-12; standardization של תאריך מפוצל מייצר פלט מתוקן וסטטוס.
- חסר: לפי המסמך השדה אינו חובה, אך ה-validator מתייחס להיעדר חודש כשגיאה.
- מעבר לדרישה: השדה נכלל ב-export schema בכל הלשוניות, כולל צורות הכתיבה האלטרנטיביות של MeshkeyBayt / AnasheyTzevet.
- נתיב פעיל או legacy: פעיל ב־Web/API Dataset Pipeline; קיימת גם התאמה ישנה היכן שרלוונטי.
- קוד / בדיקות: פעיל ב־Web/API Dataset Pipeline. בדיקות: tests/test_date_engine.py, tests/test_institution_report_validator.py. קיימת גם התאמת ייצוא ישנה מקבילה.
- הערות סיכון: פעיל ב־Web/API Dataset Pipeline. בדיקות: tests/test_date_engine.py, tests/test_institution_report_validator.py. קיימת גם התאמת ייצוא ישנה מקבילה.

### YomKnisa
- גיליונות: DayarimYahidim
- נדרש לפי המסמך: None
- דרישות המסמך: טווח ערכים תקינים בין 1 - 31 (תלוי חודש)
- ממומש: כן / חלקי / מעבר לדרישה
- מצב במערכת: ה-validator דורש את YomKnisa רק ב־DayarimYahidim; במנוע התאריכים קיימת בדיקת טווח יום 1-31.
- חסר: ה־export schema כולל את YomKnisa גם ב־AnasheyTzevet וב־MeshkeyBayt, בעוד שהמסמך מגביל את השדה ל־DayarimYahidim.
- מעבר לדרישה: גם מנוע הייצוא הישן כולל את הכיסוי הרחב הזה, ולכן מדובר באי־התאמת סקופ ולא בכשל ריצה.
- נתיב פעיל או legacy: פעיל ב־Web/API Dataset Pipeline; קיימת אי־התאמת סקופ מול המסמך.
- קוד / בדיקות: פעיל ב־Web/API Dataset Pipeline, אך עם אי־התאמת סקופ מול המסמך. בדיקות: tests/test_institution_report_validator.py, tests/test_date_engine.py.
- הערות סיכון: פעיל ב־Web/API Dataset Pipeline, אך עם אי־התאמת סקופ מול המסמך. בדיקות: tests/test_institution_report_validator.py, tests/test_date_engine.py.

## דרישות שממומשות במלואן
- MosadID: כיסוי יצוא, הזרקת מטא־דטה ודיווח.
- Min: מיפוי ולידציה של קוד מין.
- תאריכים: פריקה, תיקון וולידציה של תאריך מפוצל ומשולב.
- MisparZehut: ספרת ביקורת וזיהוי כפילויות בתוך גיליון ובין גיליונות.

## דרישות שממומשות חלקית
- SugMosad: נומריות ואורך מינימלי, ללא בדיקת מילון.
- MisparDiraBeMosad: טיפול נומרי ללא אכיפת חובה מלאה ב-validator.
- MisparZehut: ללא בדיקת מרשם אוכלוסין או מוסדות קשורים.
- shnatknisa: ללא כלל גיל כניסה מינימלי לפי סוג מוסד.
- YomKnisa: אי־התאמת סקופ בין המסמך לבין ה-export schema.

## דרישות שחסרות
- בדיקת מילון עבור SugMosad.
- הצלבה מול מרשם אוכלוסין עבור MisparZehut.
- בדיקה מול מוסדות קשורים עבור MisparZehut.
- גיל כניסה מינימלי לפי סוג מוסד.

## התנהגויות מעבר למסמך
- אכיפת ShemPrati ו־ShemMishpaha כחובה, אף שהמסמך מסמן אותם כלא מוגדרים או לא חובה.
- אכיפת רכיבי תאריך לידה וכניסה כחובה, אף שהמסמך מציין שהם לא חובה.
- ניקוי מזהים: הסרת מקפים, ריפוד, ספרת ביקורת, דחיית אפסים בלבד, והעברת מזהה לא תקין לדרכון.
- מנוע התאריכים: תאריך מדויק, שנת 1906, אזהרות עתידיות ותיקון שנה דו־ספרתית.
- סיכום קומפקטי של דוח עיבוד והודעות _validation_status ב־UI.

## התנהגות legacy בלבד שאינה פעילה
- `src/excel_standardization/export/export_engine.py` הוא קוד תאימות היסטורי ל־CLI ול־orchestrator.
- `src/excel_standardization/workbook_json_flow.py` מנתב את ייצוא ה־JSON של ה־CLI דרך מנוע הייצוא הישן ולא דרך הנתיב הפעיל של ה־Web/API.
- `src/excel_standardization/orchestrator.py` שומר את המתודות הישנות כבויות או מנותבות מהנתיב הפעיל של הנתונים.