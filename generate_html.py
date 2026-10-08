import pandas as pd
import math
import json
import random

# Files
file_teachers = "/Users/amin/Documents/rnfc/التوظيف2026.xlsx"
file_rnfc = "/Users/amin/Documents/rnfc/Besoin RNFC_Takwin_Wilaya_Branches(3).xlsx"

xls_teachers = pd.ExcelFile(file_teachers)
df_rnfc = pd.read_excel(file_rnfc)

# Mapping Arabic Wilayas to French
wilaya_map = {
    'أدرار': 'Adrar', 'الشلف': 'Chlef', 'الأغواط': 'Laghouat', 'أم البواقي': 'Oum El Bouaghi',
    'باتنة': 'Batna', 'بجاية': 'Béjaïa', 'بسكرة': 'Biskra', 'بشار': 'Béchar', 'البليدة': 'Blida',
    'البويرة': 'Bouira', 'تمنراست': 'Tamanrasset', 'تبسة': 'Tébessa', 'تلمسان': 'Tlemcen',
    'تيارت': 'Tiaret', 'تيزي وزو': 'Tizi Ouzou', 'الجزائر': 'Alger', 'الجلفة': 'Djelfa',
    'جيجل': 'Jijel', 'سطيف': 'Sétif', 'سعيدة': 'Saïda', 'سكيكدة': 'Skikda', 'سيدي بلعباس': 'Sidi Bel Abbès',
    'عنابة': 'Annaba', 'قالمة': 'Guelma', 'قسنطينة': 'Constantine', 'المدية': 'Médéa',
    'مستغانم': 'Mostaganem', 'المسيلة': "M'Sila", 'معسكر': 'Mascara', 'ورقلة': 'Ouargla',
    'وهران': 'Oran', 'البيض': 'El Bayadh', 'إليزي': 'Illizi', 'برج بوعريريج': 'Bordj Bou Arreridj',
    'بومرداس': 'Boumerdès', 'الطارف': 'El Tarf', 'تيندوف': 'Tindouf', 'تيسمسيلت': 'Tissemsilt',
    'الوادي': 'El Oued', 'خنشلة': 'Khenchela', 'سوق أهراس': 'Souk Ahras', 'تيبازة': 'Tipaza',
    'ميلة': 'Mila', 'عين الدفلى': 'Aïn Defla', 'النعامة': 'Naâma', 'عين تيموشنت': 'Aïn Témouchent',
    'غرداية': 'Ghardaïa', 'غليزان': 'Relizane', 'تيميمون': 'Timimoun', 'برج باجي مختار': 'Bordj Badji Mokhtar',
    'أولاد جلال': 'Ouled Djellal', 'بني عباس': 'Béni Abbès', 'إن صالح': 'Aïn Salah',
    'إن قزام': 'Aïn Guezzam', 'توقرت': 'Touggourt', 'جانت': 'Djanet', 'المغير': "El M'Ghair",
    'المنيعة': 'El Meniaa'
}

global_stats = {
    'total_teachers': 0,
    'total_rnfc': 0,
    'total_wilayas': 0,
    'wilayas_stats': [],
    'teacher_specialties': {},
    'rnfc_specialties': {}
}

def get_nav(active_page):
    data_class = "nav-link active" if active_page == 'data' else "nav-link"
    stats_class = "nav-link active" if active_page == 'stats' else "nav-link"
    
    return f"""
        <nav class="navbar">
            <a href="index.html" class="{data_class}">
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line><polyline points="10 9 9 9 8 9"></polyline></svg>
                البيانات الكاملة (جداول)
            </a>
            <a href="statistics.html" class="{stats_class}">
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2" ry="2"></rect><line x1="12" y1="8" x2="12" y2="21"></line><line x1="16" y1="12" x2="16" y2="21"></line><line x1="8" y1="16" x2="8" y2="21"></line></svg>
                الإحصائيات والرسوم البيانية
            </a>
        </nav>
    """

common_css = """
        :root {
            --primary: #2c3e50;
            --secondary: #34495e;
            --accent: #3498db;
            --rnfc-bg: #e8f4f8;
            --background: #f4f6f8;
            --surface: #ffffff;
            --border: #e2e8f0;
            --text: #2d3748;
        }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: var(--background); color: var(--text); margin: 0; padding: 40px 20px; }
        .container { max-width: 1400px; margin: 0 auto; }
        h1 { text-align: center; color: var(--primary); margin-bottom: 20px; font-size: 2.5rem; text-shadow: 1px 1px 2px rgba(0,0,0,0.1); }
        
        .navbar {
            display: flex;
            justify-content: center;
            gap: 20px;
            margin-bottom: 40px;
            background: var(--surface);
            padding: 15px;
            border-radius: 12px;
            box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);
        }
        .nav-link {
            display: flex;
            align-items: center;
            gap: 8px;
            padding: 12px 24px;
            background-color: transparent;
            color: var(--secondary);
            border-radius: 8px;
            text-decoration: none;
            font-weight: bold;
            font-size: 1.1rem;
            transition: all 0.3s ease;
            border: 2px solid transparent;
        }
        .nav-link:hover {
            background-color: #f1f5f9;
            color: var(--accent);
        }
        .nav-link.active {
            background-color: var(--accent);
            color: white;
            box-shadow: 0 4px 6px rgba(52, 152, 219, 0.3);
        }
        .nav-link.active svg { stroke: white; }
        .nav-link svg { stroke: currentColor; }
        
        /* Responsive Design */
        @media (max-width: 768px) {
            body { padding: 15px 10px; }
            h1 { font-size: 1.8rem; }
            .navbar { flex-direction: column; gap: 10px; padding: 10px; }
            .nav-link { justify-content: center; text-align: center; }
            .summary { flex-direction: column; gap: 15px; padding: 15px; }
            .summary div { border-right: none !important; padding-right: 0 !important; border-bottom: 1px dashed #bdc3c7; padding-bottom: 10px; margin-right: 0 !important; }
            .summary div:last-child { border-bottom: none; padding-bottom: 0; }
            .export-btn { width: 100%; justify-content: center; font-size: 1rem; }
        }
"""

html = f"""
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>جدول التوظيف 2026 - سلك الأساتذة واحتياجات RNFC</title>
    <style>
        {common_css}
        .container {{ max-width: 100%; }}
        .export-container {{ text-align: center; margin-bottom: 40px; }}
        .export-btn {{ background-color: #27ae60; color: white; padding: 12px 28px; font-size: 1.2rem; border: none; border-radius: 8px; cursor: pointer; box-shadow: 0 4px 6px rgba(0,0,0,0.1); transition: all 0.3s ease; font-weight: bold; display: inline-flex; align-items: center; gap: 10px; }}
        .export-btn:hover {{ background-color: #219653; transform: translateY(-2px); }}
        .wilaya-section {{ background: var(--surface); border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); margin-bottom: 40px; padding: 25px; overflow-x: auto; }}
        h2.wilaya-title {{ color: var(--primary); border-bottom: 3px solid var(--accent); padding-bottom: 10px; margin-top: 0; margin-bottom: 20px; display: inline-block; }}
        table {{ width: 100%; border-collapse: collapse; text-align: right; margin-bottom: 20px; min-width: 1200px; }}
        th, td {{ padding: 10px 8px; border-bottom: 1px solid var(--border); vertical-align: top; font-size: 0.95rem; }}
        th {{ background-color: var(--secondary); color: white; font-weight: 600; font-size: 1rem; }}
        .rnfc-header {{ background-color: #2980b9; }}
        .rnfc-cell {{ background-color: var(--rnfc-bg); border-right: 2px solid #bdc3c7; }}
        tr:last-child td {{ border-bottom: none; }}
        tbody tr:hover {{ background-color: #f1f5f9; }}
        .summary {{ display: flex; justify-content: flex-start; gap: 40px; font-size: 1.1rem; font-weight: bold; background-color: #e2e8f0; padding: 15px 25px; border-radius: 8px; color: var(--primary); margin-bottom: 10px; }}
        .summary span {{ color: #e74c3c; font-size: 1.2rem; margin-right: 10px; }}
        .rnfc-start {{ border-right: 3px solid #2980b9 !important; }}
    </style>
</head>
<body>
    <div class="container">
        <h1>جدول التوظيف 2026 - سلك الأساتذة واحتياجات RNFC</h1>
        {get_nav('data')}
        <div class="export-container">
            <button onclick="exportToExcel()" class="export-btn">
                <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line><polyline points="10 9 9 9 8 9"></polyline></svg>
                تحميل البيانات كملف إكسل
            </button>
        </div>
"""

for sheet_name in xls_teachers.sheet_names:
    df_t = pd.read_excel(xls_teachers, sheet_name=sheet_name)
    wilaya_ar = sheet_name
    directorate = f"مديرية التكوين لولاية {wilaya_ar}"
    
    teachers_data = []
    institution = ""
    rank = ""
    for i in range(5, len(df_t)):
        row = df_t.iloc[i]
        col_inst = str(row.iloc[0])
        col_rank = str(row.iloc[1])
        col_spec = str(row.iloc[2])
        col_num  = str(row.iloc[3])
        
        if col_inst != "nan" and col_inst.strip() != "":
            institution = col_inst.strip()
            
        if col_rank != "nan" and col_rank.strip() != "":
            rank = col_rank.strip()
            
        if (col_num != "nan" and col_num.strip() != "") or (col_spec != "nan" and col_spec.strip() != ""):
            teachers_data.append({
                "dir": directorate,
                "inst": institution,
                "rank": rank,
                "spec": col_spec if col_spec != 'nan' else '',
                "num": col_num if col_num != 'nan' else ''
            })

    wilaya_fr = wilaya_map.get(wilaya_ar, None)
    rnfc_data = []
    if wilaya_fr:
        rnfc_rows = df_rnfc[df_rnfc.iloc[:, 1] == wilaya_fr]
        for _, r in rnfc_rows.iterrows():
            code = str(r.iloc[2])
            spec = str(r.iloc[3])
            num = str(r.iloc[4])
            
            if code != "nan" or spec != "nan":
                rnfc_data.append({
                    "code": code if code != "nan" else "",
                    "spec": spec if spec != "nan" else "",
                    "num": num if num != "nan" else ""
                })
    
    if len(teachers_data) == 0 and len(rnfc_data) == 0:
        continue
        
    total_teachers = 0
    wilaya_t_specs = {}
    for t in teachers_data:
        try:
            val = int(float(t['num']))
            total_teachers += val
            if t['spec']:
                global_stats['teacher_specialties'][t['spec']] = global_stats['teacher_specialties'].get(t['spec'], 0) + val
                wilaya_t_specs[t['spec']] = wilaya_t_specs.get(t['spec'], 0) + val
        except ValueError:
            pass

    total_rnfc = 0
    wilaya_r_specs = {}
    for r in rnfc_data:
        try:
            val = int(float(r['num']))
            total_rnfc += val
            if r['spec']:
                global_stats['rnfc_specialties'][r['spec']] = global_stats['rnfc_specialties'].get(r['spec'], 0) + val
                wilaya_r_specs[r['spec']] = wilaya_r_specs.get(r['spec'], 0) + val
        except ValueError:
            pass
            
    global_stats['total_teachers'] += total_teachers
    global_stats['total_rnfc'] += total_rnfc
    global_stats['total_wilayas'] += 1
    
    # Store Wilaya specific data
    global_stats['wilayas_stats'].append({
        'wilaya': wilaya_ar,
        't_sum': total_teachers,
        'r_sum': total_rnfc,
        't_specs': len(wilaya_t_specs),
        'r_specs': len(wilaya_r_specs),
        'top_t': sorted(wilaya_t_specs.items(), key=lambda x: x[1], reverse=True)[:5],
        'top_r': sorted(wilaya_r_specs.items(), key=lambda x: x[1], reverse=True)[:5]
    })

    max_len = max(len(teachers_data), len(rnfc_data))
    
    rows_html = ""
    for i in range(max_len):
        t = teachers_data[i] if i < len(teachers_data) else {"dir": "", "inst": "", "rank": "", "spec": "", "num": ""}
        r = rnfc_data[i] if i < len(rnfc_data) else {"code": "", "spec": "", "num": ""}
        
        rows_html += f"                    <tr>\n"
        rows_html += f"                        <td>{t['dir']}</td>\n"
        rows_html += f"                        <td>{t['inst']}</td>\n"
        rows_html += f"                        <td>{t['rank']}</td>\n"
        rows_html += f"                        <td>{t['spec']}</td>\n"
        rows_html += f"                        <td>{t['num']}</td>\n"
        
        if r['spec'] != "" or r['num'] != "":
            rnfc_inst = t['inst'] if t['inst'] != "" else ""
            rows_html += f"                        <td class='rnfc-cell rnfc-start'>{r['spec']}</td>\n"
            rows_html += f"                        <td class='rnfc-cell'>{r['num']}</td>\n"
            rows_html += f"                        <td class='rnfc-cell'>{directorate}</td>\n"
            rows_html += f"                        <td class='rnfc-cell'>{rnfc_inst}</td>\n"
            rows_html += f"                        <td class='rnfc-cell'></td>\n"
            rows_html += f"                        <td class='rnfc-cell'></td>\n"
            rows_html += f"                        <td class='rnfc-cell'></td>\n"
        else:
            rows_html += f"                        <td class='rnfc-cell rnfc-start'></td>\n"
            rows_html += f"                        <td class='rnfc-cell'></td>\n"
            rows_html += f"                        <td class='rnfc-cell'></td>\n"
            rows_html += f"                        <td class='rnfc-cell'></td>\n"
            rows_html += f"                        <td class='rnfc-cell'></td>\n"
            rows_html += f"                        <td class='rnfc-cell'></td>\n"
            rows_html += f"                        <td class='rnfc-cell'></td>\n"
            
        rows_html += f"                    </tr>\n"

    forty_percent = total_teachers * 0.4
    forty_percent_rounded = int(round(forty_percent))
    
    html += f"""
        <div class="wilaya-section">
            <h2 class="wilaya-title">الولاية: {wilaya_ar}</h2>
            <table>
                <thead>
                    <tr>
                        <th colspan="5">بيانات التوظيف - سلك الأساتذة</th>
                        <th colspan="7" class="rnfc-header rnfc-start">احتياجات RNFC</th>
                    </tr>
                    <tr>
                        <th>مديرية التكوين</th>
                        <th>مؤسسة تكوينية</th>
                        <th>الرتبة</th>
                        <th>التخصص</th>
                        <th>العدد</th>
                        
                        <th class="rnfc-header rnfc-start">Nom spécialité</th>
                        <th class="rnfc-header">Effectif confirmé / Inscriptions</th>
                        <th class="rnfc-header">مديرية التكوين</th>
                        <th class="rnfc-header">مؤسسة تكوينية</th>
                        <th class="rnfc-header">الرتبة</th>
                        <th class="rnfc-header">التخصص</th>
                        <th class="rnfc-header">العدد</th>
                    </tr>
                </thead>
                <tbody>
{rows_html}
                </tbody>
            </table>
            <div class="summary">
                <div>مجموع المناصب (سلك الأساتذة): <span>{total_teachers}</span></div>
                <div>حصة 40% (أساتذة): <span>{forty_percent:g}</span> <small style='color:#555; font-size:0.9rem;'>(حوالي {forty_percent_rounded} منصب)</small></div>
                <div style="margin-right: auto; padding-right: 20px; border-right: 2px solid #bdc3c7;">مجموع احتياجات RNFC: <span>{total_rnfc}</span></div>
            </div>
        </div>
    """

html += """
    </div>

    <script src="https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.min.js"></script>
    <script>
        function exportToExcel() {
            var wb = XLSX.utils.book_new();
            var sections = document.querySelectorAll('.wilaya-section');
            var combinedHtml = "<table><tbody>";
            
            sections.forEach(function(section) {
                var title = section.querySelector('.wilaya-title').innerText;
                var summaryDivs = section.querySelectorAll('.summary div');
                var rows = section.querySelectorAll('table tr');
                
                combinedHtml += "<tr><th colspan='12'>" + title + "</th></tr>";
                rows.forEach(function(row) {
                    combinedHtml += row.outerHTML;
                });
                
                var sum1 = summaryDivs[0].innerText;
                var sum2 = summaryDivs[1].innerText;
                var sum3 = summaryDivs[2].innerText;
                
                combinedHtml += "<tr>";
                combinedHtml += "<td colspan='5'>" + sum1 + "  |  " + sum2 + "</td>";
                combinedHtml += "<td colspan='7'>" + sum3 + "</td>";
                combinedHtml += "</tr>";
                combinedHtml += "<tr><td colspan='12'></td></tr>";
            });
            
            combinedHtml += "</tbody></table>";
            var tempDiv = document.createElement('div');
            tempDiv.innerHTML = combinedHtml;
            var finalTable = tempDiv.querySelector('table');
            var ws = XLSX.utils.table_to_sheet(finalTable);
            
            ws['!cols'] = [
                {wch: 25}, {wch: 40}, {wch: 20}, {wch: 25}, {wch: 10}, 
                {wch: 35}, {wch: 15}, {wch: 25}, {wch: 40}, {wch: 20}, {wch: 25}, {wch: 12}
            ];
            
            for (var cell in ws) {
                if (cell[0] === '!') continue;
                var cellValue = ws[cell].v;
                if (!ws[cell].s) ws[cell].s = {};
                ws[cell].s.alignment = { vertical: 'center', horizontal: 'center', wrapText: true };
                ws[cell].s.font = { name: 'Arial', sz: 11 };
                
                if (typeof cellValue === 'string') {
                    if (cellValue.includes('الولاية:')) {
                        ws[cell].s.fill = { fgColor: { rgb: "E2E8F0" } };
                        ws[cell].s.font = { bold: true, sz: 14, color: { rgb: "2C3E50" } };
                    }
                    else if (cellValue === 'بيانات التوظيف - سلك الأساتذة') {
                        ws[cell].s.fill = { fgColor: { rgb: "34495E" } };
                        ws[cell].s.font = { bold: true, color: { rgb: "FFFFFF" } };
                    }
                    else if (cellValue === 'احتياجات RNFC') {
                        ws[cell].s.fill = { fgColor: { rgb: "2980B9" } };
                        ws[cell].s.font = { bold: true, color: { rgb: "FFFFFF" } };
                    }
                    else if (['مديرية التكوين', 'مؤسسة تكوينية', 'الرتبة', 'التخصص', 'العدد', 'Nom spécialité', 'Effectif confirmé / Inscriptions'].includes(cellValue)) {
                        ws[cell].s.fill = { fgColor: { rgb: "7F8C8D" } };
                        ws[cell].s.font = { bold: true, color: { rgb: "FFFFFF" } };
                    }
                    else if (cellValue.includes('مجموع المناصب')) {
                        ws[cell].s.fill = { fgColor: { rgb: "F1C40F" } };
                        ws[cell].s.font = { bold: true, color: { rgb: "2C3E50" } };
                    }
                    else if (cellValue.includes('مجموع احتياجات')) {
                        ws[cell].s.fill = { fgColor: { rgb: "F1C40F" } };
                        ws[cell].s.font = { bold: true, color: { rgb: "2C3E50" } };
                    }
                }
            }
            XLSX.utils.book_append_sheet(wb, ws, "جميع_الولايات");
            XLSX.writeFile(wb, "البيانات_الكاملة_للتوظيف_والارنفاسي_2026.xlsx");
        }
    </script>
</body>
</html>
"""

with open('/Users/amin/Documents/rnfc/index.html', 'w', encoding='utf-8') as f:
    f.write(html)


# ---------------- STATISTICS PAGE ----------------


all_wilayas_sorted = sorted(global_stats['wilayas_stats'], key=lambda x: x['r_sum'], reverse=True)
all_t_specs = sorted(global_stats['teacher_specialties'].items(), key=lambda x: x[1], reverse=True)
all_r_specs = sorted(global_stats['rnfc_specialties'].items(), key=lambda x: x[1], reverse=True)

# Wilayas JSON for dynamic interactivity
wilayas_json = json.dumps({
    w['wilaya']: {
        't_sum': w['t_sum'],
        'r_sum': w['r_sum'],
        'top_t': w['top_t'],
        'top_r': w['top_r']
    } for w in all_wilayas_sorted
}, ensure_ascii=False)

# Chart Data JSON
top_w_t = sorted(global_stats['wilayas_stats'], key=lambda x: x['t_sum'], reverse=True)[:10]
top_w_r = sorted(global_stats['wilayas_stats'], key=lambda x: x['r_sum'], reverse=True)[:10]

chart_data = {
    'top_w_t_labels': [w['wilaya'] for w in top_w_t],
    'top_w_t_data': [w['t_sum'] for w in top_w_t],
    'top_w_r_labels': [w['wilaya'] for w in top_w_r],
    'top_w_r_data': [w['r_sum'] for w in top_w_r],
    'top_s_t_labels': [s[0] for s in all_t_specs[:10]],
    'top_s_t_data': [s[1] for s in all_t_specs[:10]],
    'top_s_r_labels': [s[0] for s in all_r_specs[:10]],
    'top_s_r_data': [s[1] for s in all_r_specs[:10]]
}
chart_data_json = json.dumps(chart_data, ensure_ascii=False)

def generate_table_rows(specs_list, color):
    rows = ""
    for k, v in specs_list:
        rows += f"<tr><td>{k}</td><td><span class='badge' style='background-color: {color}'>{v}</span></td></tr>"
    return rows

stats_html = f'''
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>الإحصائيات العامة - 2026</title>
    <!-- DataTables CSS -->
    <link rel="stylesheet" href="https://cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css">
    <!-- Chart.js -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        {common_css}
        
        .grid-cards {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin-bottom: 40px;
        }}
        .card {{
            background: var(--surface);
            border-radius: 12px;
            padding: 25px;
            box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);
            text-align: center;
            border-bottom: 4px solid var(--accent);
            transition: transform 0.3s ease;
        }}
        .card:hover {{ transform: translateY(-5px); }}
        .card h3 {{ margin: 0; color: #7f8c8d; font-size: 1.1rem; }}
        .card .value {{ font-size: 2.5rem; font-weight: bold; color: var(--primary); margin-top: 10px; }}
        
        .charts-grid {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 25px;
            margin-bottom: 40px;
        }}
        .chart-box {{
            background: var(--surface);
            border-radius: 12px;
            padding: 25px;
            box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);
            position: relative;
        }}
        .chart-box h3 {{ color: var(--secondary); border-bottom: 2px solid var(--border); padding-bottom: 10px; margin-top: 0; margin-bottom: 20px; text-align: center; font-size: 1.4rem; }}
        
        .badge {{ background-color: var(--accent); color: white; padding: 6px 12px; border-radius: 20px; font-size: 1rem; font-weight: bold; }}
        
        /* DataTables Custom Styling */
        table.dataTable {{ width: 100% !important; border-collapse: collapse !important; }}
        table.dataTable thead th {{ background-color: var(--secondary); color: white; border-bottom: none; text-align: right; padding: 12px; }}
        table.dataTable tbody td {{ padding: 12px; border-bottom: 1px solid var(--border); }}
        table.dataTable tbody tr:hover {{ background-color: #f1f5f9; }}
        .dataTables_wrapper .dataTables_filter input {{ border: 1px solid var(--border); border-radius: 6px; padding: 6px 10px; margin-right: 10px; outline: none; }}
        .dataTables_wrapper .dataTables_filter input:focus {{ border-color: var(--accent); }}
        
        /* Interactive Wilaya Section */
        .interactive-wilaya {{ background: linear-gradient(135deg, #2c3e50, #3498db); padding: 40px; border-radius: 16px; color: white; margin-bottom: 40px; box-shadow: 0 10px 15px -3px rgba(0,0,0,0.2); }}
        .interactive-wilaya h2 {{ margin-top: 0; text-align: center; font-size: 2rem; margin-bottom: 20px; }}
        .wilaya-select-container {{ text-align: center; margin-bottom: 30px; }}
        .wilaya-select {{ padding: 12px 20px; font-size: 1.2rem; border-radius: 8px; border: 2px solid white; background: rgba(255,255,255,0.2); color: white; outline: none; cursor: pointer; font-family: inherit; font-weight: bold; }}
        .wilaya-select option {{ background: white; color: var(--text); }}
        
        .wilaya-details-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }}
        .w-card {{ background: rgba(255,255,255,0.1); backdrop-filter: blur(10px); padding: 20px; border-radius: 12px; border: 1px solid rgba(255,255,255,0.2); }}
        .w-card h4 {{ margin: 0 0 15px 0; color: #f1c40f; border-bottom: 1px solid rgba(255,255,255,0.2); padding-bottom: 10px; }}
        .w-stat {{ font-size: 2rem; font-weight: bold; margin-bottom: 20px; text-align: center; }}
        .w-list {{ list-style: none; padding: 0; margin: 0; }}
        .w-list li {{ display: flex; justify-content: space-between; padding: 8px 0; border-bottom: 1px dashed rgba(255,255,255,0.2); }}
        .w-list li:last-child {{ border-bottom: none; }}
        .w-list li span.val {{ background: #f39c12; padding: 2px 8px; border-radius: 12px; font-weight: bold; font-size: 0.9rem; }}

        @media (max-width: 900px) {{
            .charts-grid {{ grid-template-columns: 1fr; }}
            .wilaya-details-grid {{ grid-template-columns: 1fr; }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1>الإحصائيات العامة الشاملة</h1>
        {get_nav('stats')}
        
        <div class="grid-cards">
            <div class="card" style="border-bottom-color: var(--accent);">
                <h3>عدد الولايات المشمولة</h3>
                <div class="value">{global_stats['total_wilayas']}</div>
            </div>
            <div class="card" style="border-bottom-color: #27ae60;">
                <h3>إجمالي مناصب الأساتذة</h3>
                <div class="value">{global_stats['total_teachers']}</div>
            </div>
            <div class="card" style="border-bottom-color: #f39c12;">
                <h3>إجمالي احتياجات RNFC</h3>
                <div class="value">{global_stats['total_rnfc']}</div>
            </div>
        </div>
        
        <!-- NEW CHARTS SECTION -->
        <h2 style="color: var(--primary); text-align: center; margin-top: 40px; margin-bottom: 20px; border-bottom: 3px solid var(--accent); padding-bottom: 10px; display: inline-block;">الرسوم البيانية التوضيحية</h2>
        
        <div class="charts-grid">
            <div class="chart-box">
                <h3>أكثر 10 ولايات طلباً للأساتذة</h3>
                <canvas id="chartWT"></canvas>
            </div>
            <div class="chart-box">
                <h3>أكثر 10 ولايات طلباً للـ RNFC</h3>
                <canvas id="chartWR"></canvas>
            </div>
        </div>

        <div class="charts-grid">
            <div class="chart-box" style="height: 400px; display: flex; flex-direction: column; align-items: center;">
                <h3 style="width: 100%;">أكثر 10 تخصصات مطلوبة (أساتذة)</h3>
                <div style="flex-grow: 1; width: 100%; position: relative;">
                    <canvas id="chartST"></canvas>
                </div>
            </div>
            <div class="chart-box" style="height: 400px; display: flex; flex-direction: column; align-items: center;">
                <h3 style="width: 100%;">أكثر 10 تخصصات مطلوبة (RNFC)</h3>
                <div style="flex-grow: 1; width: 100%; position: relative;">
                    <canvas id="chartSR"></canvas>
                </div>
            </div>
        </div>

        <!-- NEW: Interactive Wilaya Profiles -->
        <div class="interactive-wilaya">
            <h2>استعراض التخصصات المطلوبة حسب الولاية</h2>
            <div class="wilaya-select-container">
                <select class="wilaya-select" id="wilayaSelector" onchange="updateWilayaDetails()">
                    <option value="" disabled selected>-- الرجاء اختيار ولاية --</option>
                    {"".join([f"<option value='{w['wilaya']}'>{w['wilaya']}</option>" for w in all_wilayas_sorted])}
                </select>
            </div>
            
            <div class="wilaya-details-grid" id="wilayaDetails" style="display: none;">
                <div class="w-card">
                    <h4>احتياجات الأساتذة في هذه الولاية</h4>
                    <div class="w-stat" id="wTeachersTotal">0</div>
                    <ul class="w-list" id="wTeachersList"></ul>
                </div>
                <div class="w-card">
                    <h4>احتياجات RNFC في هذه الولاية</h4>
                    <div class="w-stat" id="wRnfcTotal">0</div>
                    <ul class="w-list" id="wRnfcList"></ul>
                </div>
            </div>
        </div>
        
        <!-- Detailed Searchable Tables for All Specialties -->
        <div class="charts-grid">
            <div class="chart-box">
                <h3>دليل جميع تخصصات الأساتذة</h3>
                <p style="color: #7f8c8d; text-align: center; margin-top: -10px; margin-bottom: 20px;">يمكنك البحث عن أي تخصص أو ترتيب الجدول</p>
                <table class="stats-table" id="teachersTable">
                    <thead><tr><th>التخصص</th><th>العدد</th></tr></thead>
                    <tbody>
                        {generate_table_rows(all_t_specs, '#27ae60')}
                    </tbody>
                </table>
            </div>
            
            <div class="chart-box">
                <h3>دليل جميع تخصصات RNFC</h3>
                <p style="color: #7f8c8d; text-align: center; margin-top: -10px; margin-bottom: 20px;">يمكنك البحث عن أي تخصص أو ترتيب الجدول</p>
                <table class="stats-table" id="rnfcTable">
                    <thead><tr><th>التخصص (Nom spécialité)</th><th>العدد (Effectif)</th></tr></thead>
                    <tbody>
                        {generate_table_rows(all_r_specs, '#f39c12')}
                    </tbody>
                </table>
            </div>
        </div>

        <!-- Full Wilayas Detailed Table -->
        <div class="chart-box" style="margin-bottom: 40px;">
            <h3>الجدول الشامل لجميع الولايات</h3>
            <table class="stats-table" id="wilayasTable">
                <thead>
                    <tr>
                        <th>الولاية</th>
                        <th>عدد مناصب الأساتذة المفتوحة</th>
                        <th>حصة 40% (أساتذة)</th>
                        <th>عدد التخصصات المفتوحة (أساتذة)</th>
                        <th>إجمالي احتياجات RNFC</th>
                        <th>عدد التخصصات المطلوبة (RNFC)</th>
                    </tr>
                </thead>
                <tbody>
'''

for w in all_wilayas_sorted:
    forty = int(round(w['t_sum'] * 0.4))
    stats_html += f'''
                    <tr>
                        <td style="font-weight: bold; color: var(--primary);">{w['wilaya']}</td>
                        <td>{w['t_sum']}</td>
                        <td>{forty}</td>
                        <td>{w['t_specs']}</td>
                        <td>{w['r_sum']}</td>
                        <td>{w['r_specs']}</td>
                    </tr>
    '''

stats_html += f'''
                </tbody>
            </table>
        </div>

    </div>

    <!-- jQuery & DataTables JS -->
    <script src="https://code.jquery.com/jquery-3.7.0.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
    
    <script>
        // Data for Charts
        const chartData = {chart_data_json};

        // Initialize Charts
        Chart.defaults.font.family = "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif";
        Chart.defaults.color = '#7f8c8d';

        // Bar Chart - Top Wilayas Teachers
        new Chart(document.getElementById('chartWT'), {{
            type: 'bar',
            data: {{
                labels: chartData.top_w_t_labels,
                datasets: [{{
                    label: 'مناصب الأساتذة',
                    data: chartData.top_w_t_data,
                    backgroundColor: '#27ae60',
                    borderRadius: 6
                }}]
            }},
            options: {{ responsive: true, plugins: {{ legend: {{ display: false }} }} }}
        }});

        // Bar Chart - Top Wilayas RNFC
        new Chart(document.getElementById('chartWR'), {{
            type: 'bar',
            data: {{
                labels: chartData.top_w_r_labels,
                datasets: [{{
                    label: 'احتياجات RNFC',
                    data: chartData.top_w_r_data,
                    backgroundColor: '#f39c12',
                    borderRadius: 6
                }}]
            }},
            options: {{ responsive: true, plugins: {{ legend: {{ display: false }} }} }}
        }});

        // Pie Chart - Top Specialties Teachers
        new Chart(document.getElementById('chartST'), {{
            type: 'doughnut',
            data: {{
                labels: chartData.top_s_t_labels,
                datasets: [{{
                    data: chartData.top_s_t_data,
                    backgroundColor: ['#1abc9c', '#2ecc71', '#3498db', '#9b59b6', '#34495e', '#f1c40f', '#e67e22', '#e74c3c', '#ecf0f1', '#95a5a6']
                }}]
            }},
            options: {{ responsive: true, maintainAspectRatio: false }}
        }});

        // Pie Chart - Top Specialties RNFC
        new Chart(document.getElementById('chartSR'), {{
            type: 'doughnut',
            data: {{
                labels: chartData.top_s_r_labels,
                datasets: [{{
                    data: chartData.top_s_r_data,
                    backgroundColor: ['#1abc9c', '#2ecc71', '#3498db', '#9b59b6', '#34495e', '#f1c40f', '#e67e22', '#e74c3c', '#ecf0f1', '#95a5a6']
                }}]
            }},
            options: {{ responsive: true, maintainAspectRatio: false }}
        }});

        // Wilayas JSON Data for Interactivity
        const wilayasData = {wilayas_json};
        
        function updateWilayaDetails() {{
            const selected = document.getElementById('wilayaSelector').value;
            if (!selected) return;
            
            const data = wilayasData[selected];
            document.getElementById('wilayaDetails').style.display = 'grid';
            
            document.getElementById('wTeachersTotal').innerText = data.t_sum + ' منصب';
            document.getElementById('wRnfcTotal').innerText = data.r_sum + ' مقعد';
            
            // Render Teachers list
            const tList = document.getElementById('wTeachersList');
            tList.innerHTML = '';
            if (data.top_t.length === 0) tList.innerHTML = '<li>لا توجد تخصصات مطلوبة</li>';
            data.top_t.forEach(item => {{
                tList.innerHTML += `<li><span>${{item[0]}}</span> <span class="val">${{item[1]}}</span></li>`;
            }});
            
            // Render RNFC list
            const rList = document.getElementById('wRnfcList');
            rList.innerHTML = '';
            if (data.top_r.length === 0) rList.innerHTML = '<li>لا توجد تخصصات مطلوبة</li>';
            data.top_r.forEach(item => {{
                rList.innerHTML += `<li><span>${{item[0]}}</span> <span class="val">${{item[1]}}</span></li>`;
            }});
        }}

        $(document).ready(function() {{
            // Initialize Professional DataTables
            const tableOptions = {{
                "language": {{
                    "search": "بحث:",
                    "lengthMenu": "عرض _MENU_ مدخلات",
                    "info": "عرض _START_ إلى _END_ من أصل _TOTAL_ تخصص",
                    "infoEmpty": "لا توجد بيانات متاحة",
                    "infoFiltered": "(تصفية من أصل _MAX_ إجمالي)",
                    "paginate": {{
                        "first": "الأول",
                        "last": "الأخير",
                        "next": "التالي",
                        "previous": "السابق"
                    }},
                    "zeroRecords": "لم يتم العثور على أي نتائج مطابقة"
                }},
                "pageLength": 10,
                "order": [[ 1, "desc" ]] // Sort by numbers by default
            }};
            
            $('#teachersTable').DataTable(tableOptions);
            $('#rnfcTable').DataTable(tableOptions);
            
            // Wilayas table default sorting by RNFC total
            const wTableOptions = Object.assign({{}}, tableOptions);
            wTableOptions.order = [[ 4, "desc" ]];
            wTableOptions.pageLength = 20;
            $('#wilayasTable').DataTable(wTableOptions);
        }});
    </script>
</body>
</html>
'''

with open('/Users/amin/Documents/rnfc/statistics.html', 'w', encoding='utf-8') as f:
    f.write(stats_html)

print("Main HTML and Statistics HTML generated successfully.")
