# התאמות בנק + התאמות ספקים (גיול חובות)

אפליקציית Streamlit מתקדמת לניהול התאמות בנק והתאמות ספקים עם שליחת מיילים אוטומטית.

## תכונות

### התאמות בנק (כללים 1-12)
- **כלל 1**: התאמת OV/RC 1:1 (תאריך + סכום)
- **כלל 2**: הוראות קבע (קודים 469/515)
- **כלל 3**: העברות במקבץ-נט' (קוד 485)
- **כלל 4**: שיקים ספקים (קוד 493) עם טולרנס
- **כללים 5-10**: כללים נוספים לפי הגדרות

### התאמות ספקים (גיול חובות)
- **100% התאמה**: זיהוי אוטומטי של ספקים עם חוב מצטבר בין -2 ל-2 ש"ח
- **80% התאמה**: זיהוי ספקים עם יתרה 0 בין השורות
- **העברות חסרות**: זיהוי העברות ללא חשבונית + טיוטת מיילים אוטומטית

### שליחת מיילים
- שליחה אוטומטית/המונית לכל הספקים
- שליחה ידנית למייל ספציפי
- מעקב אחר הצלחות וכשלונות
- תיעוד שליחות במסד נתונים

### מסד נתונים
- SQLite מובנה לשמירת כל הנתונים
- טבלאות: ספקים, מיילים, מיפויים, הגדרות, לוג מיילים
- מיגרציה אוטומטית מקובץ JSON ישן

## דרישות מערכת

- Python 3.11+
- פקודת `pip` מותקנת

## התקנה מקומית

### 1. שכפול הפרויקט
```bash
git clone https://github.com/bdnhost/hatamotbankmimush.git
cd hatamotbankmimush
```

### 2. יצירת סביבה וירטואלית
```bash
python -m venv venv
source venv/bin/activate  # Linux/Mac
# או
venv\Scripts\activate  # Windows
```

### 3. התקנת תלויות
```bash
pip install -r requirements.txt
```

### 4. הרצת האפליקציה
```bash
streamlit run streamlit_app.py
```

האפליקציה תפתח בדפדפן בכתובת: `http://localhost:8501`

## פריסה ב-Cloud

### אופציה 1: Streamlit Community Cloud (מומלץ)

1. **הכנה**:
   - צור חשבון ב-[Streamlit Community Cloud](https://streamlit.io/cloud)
   - חבר את הריפוזיטורי שלך ב-GitHub

2. **פריסה**:
   - לחץ על "New app"
   - בחר את הריפוזיטורי: `bdnhost/hatamotbankmimush`
   - בחר את הענף: `claude/improve-accounting-reconciliation-019j4QXPp244zG2Eeo2C6VBd`
   - הקובץ הראשי: `streamlit_app.py`
   - לחץ על "Deploy"

3. **תוצאה**:
   - האפליקציה תהיה זמינה בכתובת ייחודית
   - עדכונים אוטומטיים עם כל push ל-GitHub

### אופציה 2: Docker + Railway/Render

#### Railway:
```bash
# התקנת Railway CLI
npm i -g @railway/cli

# התחברות
railway login

# פריסה
railway up
```

#### Render:
1. צור חשבון ב-[Render](https://render.com)
2. חבר את הריפוזיטורי מ-GitHub
3. בחר "New Web Service"
4. הגדרות:
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `streamlit run streamlit_app.py --server.port=$PORT`
   - Environment: Python 3.11

### אופציה 3: Docker (לכל פלטפורמה)

```bash
# בניית התמונה
docker build -t hatamot-bank .

# הרצה
docker run -p 8501:8501 hatamot-bank
```

## שימוש באפליקציה

### 1. התאמות בנק

1. העלה קובץ Excel עם DataSheet
2. (אופציונלי) העלה קובץ עזר להעברות (כלל 3)
3. לחץ על "הרצה 1-12"
4. הורד את הקובץ המעובד

### 2. התאמות ספקים

1. העלה קובץ גיול חובות (Excel)
2. הזן שם לקוח (אופציונלי)
3. לחץ על "🔍 נתח גיול חובות"
4. בדוק את התוצאות (100%, 80%, העברות חסרות)
5. הורד את דוח ההתאמות

### 3. שליחת מיילים

#### הגדרה ראשונית:
1. העלה קובץ עזר עם מיילי ספקים
   - פורמט: `מס' ספק | שם ספק | מייל ספק`
2. הגדר פרטי SMTP:
   - שרת: `smtp.gmail.com` (לדוגמה)
   - פורט: `587`
   - מייל שולח וסיסמה

#### שליחה המונית:
- לחץ על "📨 שלח מיילים לכל הספקים"
- עקוב אחר ההתקדמות
- בדוק תוצאות (הצלחות/כשלונות)

#### שליחה ידנית:
- בחר ספק מהרשימה
- הזן מייל נמען
- ערוך את תוכן המייל (אופציונלי)
- לחץ על "שלח מייל בודד"

## הגדרות SMTP

### Gmail:
1. עבור ל-[Google Account Security](https://myaccount.google.com/security)
2. הפעל "2-Step Verification"
3. צור "App Password"
4. השתמש ב-App Password במקום הסיסמה הרגילה

פרטי שרת:
- שרת: `smtp.gmail.com`
- פורט: `587`

### Outlook/Hotmail:
- שרת: `smtp-mail.outlook.com`
- פורט: `587`

### Yahoo:
- שרת: `smtp.mail.yahoo.com`
- פורט: `587`

## מבנה הפרויקט

```
hatamotbankmimush/
├── streamlit_app.py      # אפליקציה ראשית
├── database.py           # מודול מסד נתונים SQLite
├── requirements.txt      # תלויות Python
├── Dockerfile           # קובץ Docker
├── .streamlit/
│   └── config.toml      # הגדרות Streamlit
├── .gitignore           # קבצים להתעלם מהם ב-Git
└── README.md            # תיעוד (הקובץ הזה)
```

## מסד נתונים

האפליקציה משתמשת ב-SQLite עם הטבלאות הבאות:

- **suppliers**: מידע על ספקים (מס' חשבון, שם, מייל, טלפון)
- **name_mappings**: מיפוי שמות לספקים (VLOOKUP)
- **amount_mappings**: מיפוי סכומים לספקים (VLOOKUP)
- **settings**: הגדרות אפליקציה
- **email_log**: לוג שליחות מיילים

הנתונים נשמרים ב-`app_database.db` (לא נכלל ב-Git).

## תמיכה וטיפול בבעיות

### בעיות נפוצות

**הריפוזיטורי

 לא מתעדכן**:
```bash
git pull origin claude/improve-accounting-reconciliation-019j4QXPp244zG2Eeo2C6VBd
```

**שגיאות התקנה**:
```bash
pip install --upgrade pip
pip install -r requirements.txt --force-reinstall
```

**האפליקציה לא נטענת**:
```bash
streamlit cache clear
streamlit run streamlit_app.py
```

**מיילים לא נשלחים**:
- בדוק חיבור לאינטרנט
- וודא שפרטי SMTP נכונים
- בדוק שאין חסימת firewall
- השתמש ב-App Password (לא סיסמה רגילה)

## רישיון

כל הזכויות שמורות © 2025

## יוצר

פותח על ידי Claude AI עבור bdnhost
