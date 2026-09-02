import streamlit as st
import pandas as pd
import random
import io

# مكتبات تصدير الملفات
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# ---------------------------------------------------------
# إعدادات الصفحة والاتجاه العربي (RTL)
# ---------------------------------------------------------
st.set_page_config(
    page_title="نظام جدول قسم اللغة العربية",
    page_icon="📚",
    layout="wide"
)

# تطبيق تنسيقات CSS لدعم اللغة العربية والألوان
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Tajawal', sans-serif;
        direction: rtl;
        text-align: right;
    }
    
    .stSelectbox label, .stTextInput label, .stMultiselect label {
        text-align: right;
        display: block;
        font-weight: bold;
    }
    
    .course-card {
        padding: 8px;
        border-radius: 6px;
        color: #1e1e1e;
        font-weight: bold;
        box-shadow: 0 2px 4px rgba(0,0,0,0.08);
        margin-bottom: 4px;
        text-align: center;
        font-size: 13px;
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# الثوابت الخيارات المتاحة
# ---------------------------------------------------------
TIME_SLOTS = [
    "08:15 - 09:05", "08:15 - 09:30",
    "09:45 - 10:35", "09:45 - 11:00",
    "11:15 - 12:05", "11:15 - 12:30",
    "12:40 - 13:30", "12:40 - 13:55",
    "14:15 - 15:30",
    "20:10 - 21:00", "21:10 - 22:00", "22:10 - 23:00"
]

DAYS_PATTERNS = ["سبت / اثنين", "أحد / ثلاثاء", "أربعاء"]

PASTEL_COLORS = [
    "#FFD1DC", "#FFDFBA", "#FFFFBA", "#BAFFC9", 
    "#BAE1FF", "#E8AEFF", "#D4F0F0", "#CCE2CB",
    "#F6EAC2", "#FFB3BA", "#C9C9FF", "#BFFCC6"
]

# ---------------------------------------------------------
# تهيئة حالة الجلسة (Session State)
# ---------------------------------------------------------
if "schedule" not in st.session_state:
    st.session_state.schedule = []

if "doctors" not in st.session_state:
    st.session_state.doctors = {}

# ---------------------------------------------------------
# دوال إنشاء وتصدير الملفات (Excel, Word, PDF)
# ---------------------------------------------------------

# 1. إنشاء ملف Excel منسق
def create_excel_schedule(schedule, time_slots, days_patterns):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "الجدول الدراسي"
    ws.views.sheetView[0].rightToLeft = True  # اتجاه من اليمين لليسامر
    
    # عنوان رئيسي
    ws.merge_cells("A1:D1")
    ws["A1"] = "جدول قسم اللغة العربية"
    ws["A1"].font = Font(name="Calibri", size=16, bold=True, color="1F4E79")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    
    headers = ["الوقت"] + days_patterns
    ws.append([])
    ws.append(headers)
    
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    
    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=3, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    df = pd.DataFrame(schedule) if schedule else pd.DataFrame()
    
    current_row = 4
    for slot in time_slots:
        row_data = [slot]
        for day in days_patterns:
            if not df.empty:
                matches = df[(df["day"] == day) & (df["time"] == slot)]
                if not matches.empty:
                    text_items = [f"{m['course']} ({m['doctor']})" for _, m in matches.iterrows()]
                    row_data.append(" / ".join(text_items))
                else:
                    row_data.append("-")
            else:
                row_data.append("-")
        ws.append(row_data)
        
        ws.cell(row=current_row, column=1).font = Font(bold=True)
        ws.cell(row=current_row, column=1).alignment = Alignment(horizontal="center", vertical="center")
        for col_num in range(2, len(headers) + 1):
            cell = ws.cell(row=current_row, column=col_num)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        current_row += 1
        
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 5, 22)
        
    wb.save(output)
    output.seek(0)
    return output.getvalue()

# 2. إنشاء ملف Word (.docx)
def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def create_word_schedule(schedule, time_slots, days_patterns):
    doc = Document()
    
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run("جدول قسم اللغة العربية")
    run_title.font.name = "Arial"
    run_title.font.size = Pt(18)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(31, 78, 121)
    
    headers = ["الأوقات"] + days_patterns
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    hdr_cells = table.rows[0].cells
    for i, header_text in enumerate(headers):
        hdr_cells[i].text = header_text
        set_cell_background(hdr_cells[i], "1F4E79")
        for paragraph in hdr_cells[i].paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                run.font.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.font.name = "Arial"
                run.font.size = Pt(11)
                
    df = pd.DataFrame(schedule) if schedule else pd.DataFrame()
    
    for slot in time_slots:
        row_cells = table.add_row().cells
        row_cells[0].text = slot
        p0 = row_cells[0].paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if p0.runs:
            p0.runs[0].font.bold = True
            p0.runs[0].font.name = "Arial"
            
        for idx, day in enumerate(days_patterns, start=1):
            if not df.empty:
                matches = df[(df["day"] == day) & (df["time"] == slot)]
                if not matches.empty:
                    cell_text = "\n".join([f"{m['course']}\n({m['doctor']})" for _, m in matches.iterrows()])
                    row_cells[idx].text = cell_text
                    set_cell_background(row_cells[idx], "F2F4F7")
                else:
                    row_cells[idx].text = "-"
            else:
                row_cells[idx].text = "-"
                
            p = row_cells[idx].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.name = "Arial"
                run.font.size = Pt(10)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output.getvalue()

# 3. إنشاء نسق PDF طباعي تفاعلي بدعم الخطوط العربية الكامل
def create_pdf_printable_html(schedule, time_slots, days_patterns):
    df = pd.DataFrame(schedule) if schedule else pd.DataFrame()
    
    html_content = """
    <!DOCTYPE html>
    <html dir="rtl" lang="ar">
    <head>
    <meta charset="UTF-8">
    <title>جدول قسم اللغة العربية</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; direction: rtl; text-align: right; margin: 20px; }
        h1 { text-align: center; color: #1F4E79; font-size: 24px; margin-bottom: 20px; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; }
        th, td { border: 1px solid #B0C4DE; padding: 10px; text-align: center; font-size: 13px; }
        th { background-color: #1F4E79; color: white; font-size: 14px; }
        tr:nth-child(even) { background-color: #F9FBFD; }
        .course-box { background-color: #EBF1F5; border-radius: 4px; padding: 6px; font-weight: bold; margin: 2px; }
        @media print {
            .no-print { display: none; }
            body { margin: 0; }
        }
    </style>
    </head>
    <body>
    <button class="no-print" onclick="window.print()" style="padding: 10px 20px; font-size: 15px; background: #1F4E79; color: white; border: none; border-radius: 5px; cursor: pointer; margin-bottom: 15px;">🖨️ اضغط هنا للطباعة أو الحفظ كـ PDF</button>
    <h1>📚 جدول قسم اللغة العربية</h1>
    <table>
        <thead>
            <tr>
                <th>الوقت</th>
    """
    for day in days_patterns:
        html_content += f"<th>{day}</th>"
    html_content += "</tr></thead><tbody>"
    
    for slot in time_slots:
        html_content += f"<tr><td style='font-weight:bold; background-color:#F0F4F8;'>{slot}</td>"
        for day in days_patterns:
            if not df.empty:
                matches = df[(df["day"] == day) & (df["time"] == slot)]
                if not matches.empty:
                    cell_str = "<br>".join([f"<div class='course-box'>{m['course']}<br><span style='font-size:11px; font-weight:normal; color:#444;'>({m['doctor']})</span></div>" for _, m in matches.iterrows()])
                    html_content += f"<td>{cell_str}</td>"
                else:
                    html_content += "<td>-</td>"
            else:
                html_content += "<td>-</td>"
        html_content += "</tr>"
        
    html_content += """
        </tbody>
    </table>
    </body>
    </html>
    """
    return html_content.encode('utf-8')

# ---------------------------------------------------------
# الشريط الجانبي: إدخال البيانات
# ---------------------------------------------------------
st.sidebar.title("⚙️ إدارة البيانات والمدخلات")

with st.sidebar.expander("➕ إضافة مدرس ومواد", expanded=True):
    doc_name = st.text_input("اسم الدكتور / أستاذ المادة:")
    courses_input = st.text_area("المواد المكلف بها (مادة في كل سطر):")
    preferred_days = st.multiselect("الأيام المتاحة للدكتور:", DAYS_PATTERNS, default=DAYS_PATTERNS)
    break_time = st.selectbox("وقت الاستراحة المفضّل:", ["لا يوجد"] + TIME_SLOTS)

    if st.button("حفظ بيانات الدكتور"):
        if doc_name and courses_input:
            courses = [c.strip() for c in courses_input.split("\n") if c.strip()]
            color = random.choice(PASTEL_COLORS)
            st.session_state.doctors[doc_name] = {
                "courses": courses,
                "days": preferred_days,
                "break": break_time,
                "color": color
            }
            st.sidebar.success(f"تم تسجيل البيانات لـ {doc_name}")
        else:
            st.sidebar.error("يرجى إدخال الاسم والمواد بشكل صحيح.")

if st.session_state.doctors:
    with st.sidebar.expander("📋 قائمة المدرسين والخيارات"):
        for doc, info in st.session_state.doctors.items():
            st.markdown(f"**{doc}**")
            st.caption(f"المواد: {', '.join(info['courses'])}")

st.sidebar.markdown("---")
if st.sidebar.button("🚀 توليد الجدول المقترح", type="primary"):
    st.session_state.schedule = []
    used_slots = set()

    for doc, info in st.session_state.doctors.items():
        for course in info["courses"]:
            assigned = False
            for day in info["days"]:
                if assigned:
                    break
                for slot in TIME_SLOTS:
                    if slot == info["break"]:
                        continue
                    
                    key = (doc, day, slot)
                    if key not in used_slots:
                        st.session_state.schedule.append({
                            "id": len(st.session_state.schedule) + 1,
                            "doctor": doc,
                            "course": course,
                            "day": day,
                            "time": slot,
                            "color": info["color"]
                        })
                        used_slots.add(key)
                        assigned = True
                        break

    st.sidebar.success("تم توليد الجدول بنجاح!")

if st.sidebar.button("🗑️ مسح جميع البيانات"):
    st.session_state.schedule = []
    st.session_state.doctors = {}
    st.rerun()

# ---------------------------------------------------------
# الواجهة الرئيسية
# ---------------------------------------------------------
st.title("📚 نظام إدارة جدول قسم اللغة العربية")
st.write("عرض، توزيع، وتعديل الجدول الدراسي مع خيارات التصدير المباشرة.")

tab1, tab2, tab3 = st.tabs(["📊 عرض الجدول العام", "🔄 التعديل والتنقل التفاعلي", "📥 التصدير والإحصائيات"])

# التبويب الأول: شبكة عرض الجدول
with tab1:
    if not st.session_state.schedule:
        st.info("لم يتم توليد أي جدول بعد. أضف أسماء المدرسين من الشريط الجانبي ثم اضغط 'توليد الجدول المقترح'.")
    else:
        st.subheader("🗓️ الجدول الدراسي الموحد")
        df = pd.DataFrame(st.session_state.schedule)
        
        pivot_data = []
        for slot in TIME_SLOTS:
            row = {"الأوقات": slot}
            for day in DAYS_PATTERNS:
                matches = df[(df["day"] == day) & (df["time"] == slot)]
                if not matches.empty:
                    cell_content = ""
                    for _, m in matches.iterrows():
                        cell_content += f"<div class='course-card' style='background-color:{m['color']}'>{m['course']}<br><small>{m['doctor']}</small></div>"
                    row[day] = cell_content
                else:
                    row[day] = "-"
            pivot_data.append(row)

        grid_df = pd.DataFrame(pivot_data)
        st.write(grid_df.to_html(escape=False, index=False), unsafe_allow_html=True)

# التبويب الثاني: النقل والتعديل
with tab2:
    if not st.session_state.schedule:
        st.warning("لا توجد مواد في الجدول لتعديلها.")
    else:
        st.subheader("🔄 نقل مادة أو تغيير الدكتور/الموعد")
        df_sched = pd.DataFrame(st.session_state.schedule)
        course_labels = [f"#{row['id']} - {row['course']} ({row['doctor']}) | {row['day']} [{row['time']}]" for _, row in df_sched.iterrows()]
        selected_item = st.selectbox("اختر المحاضرة المراد نقلها أو تعديلها:", course_labels)

        if selected_item:
            item_id = int(selected_item.split("#")[1].split(" -")[0])
            current_data = next(item for item in st.session_state.schedule if item["id"] == item_id)

            col1, col2, col3 = st.columns(3)
            with col1:
                new_doc = st.selectbox("الدكتور المسؤول:", list(st.session_state.doctors.keys()), index=list(st.session_state.doctors.keys()).index(current_data["doctor"]))
            with col2:
                new_day = st.selectbox("الأيام:", DAYS_PATTERNS, index=DAYS_PATTERNS.index(current_data["day"]))
            with col3:
                new_time = st.selectbox("الوقت:", TIME_SLOTS, index=TIME_SLOTS.index(current_data["time"]))

            conflict = any(
                item["doctor"] == new_doc and item["day"] == new_day and item["time"] == new_time and item["id"] != item_id
                for item in st.session_state.schedule
            )

            if conflict:
                st.error("⚠️ تحذير: يوجد تعارض! الدكتور لديه محاضرة أخرى في نفس هذا الوقت واليوم.")

            if st.button("💾 حفظ التعديل والنقل"):
                for item in st.session_state.schedule:
                    if item["id"] == item_id:
                        item["doctor"] = new_doc
                        item["day"] = new_day
                        item["time"] = new_time
                        item["color"] = st.session_state.doctors[new_doc]["color"]
                st.success("تم تحديث موقع المحاضرة بنجاح!")
                st.rerun()

# التبويب الثالث: أزرار التصدير لجميع الصيغ
with tab3:
    if not st.session_state.schedule:
        st.info("قم بتوليد الجدول أولاً لتتمكن من تصديره.")
    else:
        st.subheader("📥 تصدير الجدول بجميع الصيغ المطلوبة")
        st.write("اختر الصيغة المناسبة لتنزيل الجدول مباشرة إلى جهازك:")
        
        col1, col2, col3, col4 = st.columns(4)
        
        # 1. زر تحميل Excel
        excel_bytes = create_excel_schedule(st.session_state.schedule, TIME_SLOTS, DAYS_PATTERNS)
        col1.download_button(
            label="📊 تحميل Excel (.xlsx)",
            data=excel_bytes,
            file_name="جدول_قسم_اللغة_العربية.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
        # 2. زر تحميل Word
        word_bytes = create_word_schedule(st.session_state.schedule, TIME_SLOTS, DAYS_PATTERNS)
        col2.download_button(
            label="📝 تحميل Word (.docx)",
            data=word_bytes,
            file_name="جدول_قسم_اللغة_العربية.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
        
        # 3. زر طباعة/تحميل PDF
        pdf_html_bytes = create_pdf_printable_html(st.session_state.schedule, TIME_SLOTS, DAYS_PATTERNS)
        col3.download_button(
            label="📄 طباعة / حفظ كـ PDF",
            data=pdf_html_bytes,
            file_name="جدول_قسم_اللغة_العربية.html",
            mime="text/html",
            use_container_width=True
        )
        
        # 4. زر تحميل CSV
        export_df = pd.DataFrame(st.session_state.schedule)[["id", "course", "doctor", "day", "time"]]
        export_df.columns = ["الرقم", "المادة", "الدكتور", "الأيام", "الوقت"]
        csv_bytes = export_df.to_csv(index=False).encode('utf-8-sig')
        col4.download_button(
            label="📁 تحميل CSV (.csv)",
            data=csv_bytes,
            file_name="جدول_قسم_اللغة_العربية.csv",
            mime="text/csv",
            use_container_width=True
        )
        
        st.markdown("---")
        st.subheader("📋 معاينة البيانات الخام")
        st.dataframe(export_df, use_container_width=True)
