from fastapi import FastAPI, Form
from openpyxl import Workbook, load_workbook
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates
from starlette.requests import Request
import os

app = FastAPI()
templates = Jinja2Templates(directory=".")

# تحديد مسار ملف Excel في نفس مجلد التطبيق
file_path = os.path.join(os.path.dirname(__file__), "data.xlsx")

# دالة لحفظ البيانات في ملف Excel
def save_data_to_excel(name, email):
    # إذا كان الملف موجودًا، يتم تحميله
    if os.path.exists(file_path):
        workbook = load_workbook(file_path)
        sheet = workbook.active
    else:
        # إذا لم يكن الملف موجودًا، يتم إنشاؤه وإضافة العناوين
        workbook = Workbook()
        sheet = workbook.active
        sheet['A1'] = "Name"
        sheet['B1'] = "Email"

    # إيجاد الصف التالي الفارغ وإضافة البيانات
    next_row = sheet.max_row + 1
    sheet[f'A{next_row}'] = name
    sheet[f'B{next_row}'] = email

    # حفظ الملف
    workbook.save(file_path)

@app.get("/", response_class=HTMLResponse)
async def read_form(request: Request):
    return templates.TemplateResponse("form.html", {"request": request})

@app.post("/submit/")
async def handle_form(name: str = Form(...), email: str = Form(...)):
    save_data_to_excel(name, email)
    return {"message": "تم حفظ البيانات بنجاح"}
