# כלי שקפים — Office Web Add-in לפאוורפויינט

שינוי גודל שקף בלחיצה אחת. ממשק בעברית, RTL, מינימליסטי.

---

## גדלים נתמכים

| שם | מידות | שימוש |
|---|---|---|
| מצגת רחבה | 1920 × 1080 | מצגות סטנדרטיות |
| ריבוע | 1080 × 1080 | סושיאל מדיה |
| סטורי / ורטיקל | 1080 × 1920 | אינסטגרם, טיקטוק |
| A4 | 2480 × 3508 | דוחות להדפסה |
| A3 | 3508 × 4961 | פוסטרים, הדפסה גדולה |
| מותאם אישית | לפי הזנה | כל גודל |

כל הגדלים (חוץ מריבוע ומותאם) כוללים כפתור ⟳ להחלפת כיוון.

---

## דרישות מקדימות

- חשבון Microsoft 365 (לא Office one-time purchase)
- PowerPoint Desktop (Mac)
- חיבור אינטרנט (לטעינת Office.js)

> **הערה:** הגדרה זו משתמשת בקבצים מקומיים (`file:///`) ולכן עובדת רק ב-PowerPoint Desktop ולא ב-PowerPoint Web.

---

## שלב 1 — העתקת קבצי התוסף

פתח Terminal והרץ:

```bash
mkdir -p ~/Documents/MyAddins/slide-size-addin/assets
cp index.html ~/Documents/MyAddins/slide-size-addin/
cp assets/icon-32.png ~/Documents/MyAddins/slide-size-addin/assets/
```

---

## שלב 2 — התקנת manifest.xml ב-PowerPoint

```bash
mkdir -p ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef
cp manifest.xml ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/
```

פתח PowerPoint (או הפעל מחדש אם פתוח) ← Insert → My Add-ins → **כלי שקפים**

---

## שימוש

1. לחץ על **"גודל שקף"** ב-Ribbon (טאב Home)
2. בחר גודל מרשימת הכרטיסים — השינוי מיידי
3. לחץ ⟳ על כרטיס להחלפת כיוון (לאורך ↔ לרוחב)
4. לגודל מותאם — הזן רוחב וגובה בפיקסלים ולחץ **"החל"**

> אם המצגת מכילה יותר משקף אחד, תופיע שאלת אישור לפני השינוי.

---

## פתרון בעיות

| בעיה | פתרון |
|---|---|
| הכפתור לא מופיע ב-Ribbon | ודא שה-wef path נכון ו-manifest.xml תקין; הפעל מחדש את PowerPoint |
| ה-taskpane לא נטען | ודא ש-`index.html` קיים ב-`~/Documents/MyAddins/slide-size-addin/` |
| "אין גישה לשינוי גודל" | ודא שיש לך הרשאות עריכה על הקובץ |
| התוסף לא מופיע ב-PowerPoint Web | עבודה עם `file:///` אפשרית רק ב-PowerPoint Desktop |
