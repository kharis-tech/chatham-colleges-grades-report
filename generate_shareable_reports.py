import pandas as pd
import re
from datetime import datetime
import json

def parse_subject_grades(text: str) -> dict:
    """Parse subject-grade pairs from text"""
    if pd.isna(text) or not str(text).strip() or str(text).lower() in ['nan', '-', 'n/a']:
        return {}
    
    grades = {}
    text = str(text).strip()
    
    # Common patterns for UK grades
    patterns = [
        r'([A-Za-z\s&]+?)\s*[-–]\s*([A-Z*]+\d*|Merit|Distinction|Pass|\d+)',
        r'([A-Za-z\s&]+?)\s*:\s*([A-Z*]+\d*|Merit|Distinction|Pass|\d+)',
        r'([A-Za-z\s&]+?)\s+([A-Z*]+\d*|Merit|Distinction|Pass|\d+)(?=\s|$|,)',
    ]
    
    lines = re.split(r'[,\n]', text)
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        for pattern in patterns:
            matches = re.finditer(pattern, line, re.IGNORECASE)
            for match in matches:
                subject = match.group(1).strip()
                grade = match.group(2).strip()
                
                subject = re.sub(r'\s+', ' ', subject).title()
                
                if len(subject) > 1 and grade:
                    grades[subject] = grade
                    break
    
    return grades

def grade_to_points(grade: str) -> int:
    """Convert UK grades to point system for comparison"""
    grade = grade.upper().strip()
    
    gcse_points = {'9': 9, '8': 8, '7': 7, '6': 6, '5': 5, '4': 4, '3': 3, '2': 2, '1': 1, 'U': 0}
    alevel_points = {'A*': 6, 'A': 5, 'B': 4, 'C': 3, 'D': 2, 'E': 1, 'U': 0}
    btec_points = {'D*': 4, 'DISTINCTION*': 4, 'D': 3, 'DISTINCTION': 3, 'M': 2, 'MERIT': 2, 'P': 1, 'PASS': 1}
    
    if grade in gcse_points:
        return gcse_points[grade]
    elif grade in alevel_points:
        return alevel_points[grade]
    elif grade in btec_points:
        return btec_points[grade]
    
    return -1

def compare_performance(current: str, predicted: str) -> tuple:
    """Compare current vs predicted performance"""
    if current == 'N/A' and predicted == 'N/A':
        return "No Data", "gray", "❓"
    elif current == 'N/A':
        return "Target Set", "blue", "🎯"
    elif predicted == 'N/A':
        return "Has Current Grade", "green", "📈"
    
    current_points = grade_to_points(current)
    predicted_points = grade_to_points(predicted)
    
    if current_points == -1 or predicted_points == -1:
        if current.upper() == predicted.upper():
            return "Meeting Target", "green", "✅"
        else:
            return "Different Format", "yellow", "📊"
    
    if current_points > predicted_points:
        return "Exceeding Target", "darkgreen", "🎉"
    elif current_points == predicted_points:
        return "Meeting Target", "green", "✅"
    else:
        return "Below Target", "red", "⚠️"

def generate_html_report():
    """Generate a comprehensive HTML report"""
    try:
        df = pd.read_excel('KOC Grade Tracker Form(1-52).xlsx', sheet_name='Sheet1')
        
        html_content = f"""
<!DOCTYPE html>
<html>
<head>
    <title>Colleges Grade Analysis Report</title>
    <meta charset="UTF-8">
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: #f5f5f5;
        }}
        .container {{
            max-width: 1200px;
            margin: 0 auto;
            background-color: white;
            padding: 30px;
            border-radius: 10px;
            box-shadow: 0 0 20px rgba(0,0,0,0.1);
        }}
        .header {{
            text-align: center;
            margin-bottom: 30px;
            border-bottom: 3px solid #2c3e50;
            padding-bottom: 20px;
        }}
        .header h1 {{
            color: #2c3e50;
            margin: 0;
            font-size: 2.5em;
        }}
        .header p {{
            color: #7f8c8d;
            margin: 10px 0 0 0;
            font-size: 1.1em;
        }}
        .summary-stats {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }}
        .stat-card {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 20px;
            border-radius: 10px;
            text-align: center;
        }}
        .stat-number {{
            font-size: 2.5em;
            font-weight: bold;
            margin-bottom: 5px;
        }}
        .stat-label {{
            font-size: 0.9em;
            opacity: 0.9;
        }}
        .section {{
            margin-bottom: 40px;
        }}
        .section h2 {{
            color: #2c3e50;
            border-left: 5px solid #3498db;
            padding-left: 15px;
            margin-bottom: 20px;
        }}
        .student-card {{
            background-color: #f8f9fa;
            border: 1px solid #dee2e6;
            border-radius: 8px;
            padding: 20px;
            margin-bottom: 20px;
        }}
        .student-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 15px;
            flex-wrap: wrap;
        }}
        .student-name {{
            font-size: 1.3em;
            font-weight: bold;
            color: #2c3e50;
        }}
        .student-info {{
            color: #6c757d;
            font-size: 0.9em;
        }}
        .grades-table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 10px;
        }}
        .grades-table th, .grades-table td {{
            padding: 8px 12px;
            text-align: left;
            border-bottom: 1px solid #dee2e6;
        }}
        .grades-table th {{
            background-color: #e9ecef;
            font-weight: 600;
            color: #495057;
        }}
        .status-exceeding {{ color: #28a745; font-weight: bold; }}
        .status-meeting {{ color: #28a745; }}
        .status-below {{ color: #dc3545; font-weight: bold; }}
        .status-nodata {{ color: #6c757d; }}
        .status-target {{ color: #007bff; }}
        .status-current {{ color: #17a2b8; }}
        .performance-summary {{
            background-color: #e7f3ff;
            border-left: 4px solid #007bff;
            padding: 15px;
            margin-top: 15px;
            border-radius: 0 5px 5px 0;
        }}
        .alert-attention {{
            background-color: #fff3cd;
            border: 1px solid #ffeaa7;
            border-radius: 5px;
            padding: 15px;
            margin-bottom: 20px;
        }}
        .alert-success {{
            background-color: #d4edda;
            border: 1px solid #c3e6cb;
            border-radius: 5px;
            padding: 15px;
            margin-bottom: 20px;
        }}
        .priority-high {{ border-left: 5px solid #dc3545; }}
        .priority-medium {{ border-left: 5px solid #ffc107; }}
        .priority-low {{ border-left: 5px solid #28a745; }}
        @media (max-width: 768px) {{
            .student-header {{
                flex-direction: column;
                align-items: flex-start;
            }}
            .grades-table {{
                font-size: 0.9em;
            }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🎓 Colleges Grade Analysis Report</h1>
            <p>Generated on {datetime.now().strftime('%B %d, %Y at %I:%M %p')}</p>
        </div>
"""
        
        # Process data
        students_data = []
        total_students = 0
        students_with_data = 0
        exceeding_count = 0
        meeting_count = 0
        below_count = 0
        attention_needed = []
        high_performers = []
        
        for index, row in df.iterrows():
            name = row['Full Name']
            if pd.isna(name) or name.strip() == '':
                continue
                
            total_students += 1
            
            school = row['School You Attend']
            year = row['What year are you in']
            current_text = row['Please list all the subjects you are currently taking and your current grades']
            predicted_text = row['Please list all your predicted grades for each subject']
            
            current_grades = parse_subject_grades(current_text)
            predicted_grades = parse_subject_grades(predicted_text)
            
            if current_grades or predicted_grades:
                students_with_data += 1
            
            all_subjects = set(current_grades.keys()) | set(predicted_grades.keys())
            
            student_data = {
                'name': name,
                'school': school,
                'year': year,
                'subjects': [],
                'exceeding': 0,
                'meeting': 0,
                'below': 0,
                'priority': 'low'
            }
            
            for subject in sorted(all_subjects):
                current = current_grades.get(subject, 'N/A')
                predicted = predicted_grades.get(subject, 'N/A')
                status, color, icon = compare_performance(current, predicted)
                
                student_data['subjects'].append({
                    'subject': subject,
                    'current': current,
                    'predicted': predicted,
                    'status': status,
                    'color': color,
                    'icon': icon
                })
                
                if "Exceeding" in status:
                    student_data['exceeding'] += 1
                    exceeding_count += 1
                elif "Meeting" in status:
                    student_data['meeting'] += 1
                    meeting_count += 1
                elif "Below" in status:
                    student_data['below'] += 1
                    below_count += 1
            
            # Determine priority
            if student_data['below'] >= 3:
                student_data['priority'] = 'high'
                attention_needed.append(student_data)
            elif student_data['below'] >= 1:
                student_data['priority'] = 'medium'
            
            if student_data['exceeding'] >= 2:
                high_performers.append(student_data)
            
            students_data.append(student_data)
        
        # Add summary statistics
        html_content += f"""
        <div class="summary-stats">
            <div class="stat-card">
                <div class="stat-number">{total_students}</div>
                <div class="stat-label">Total Students</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">{students_with_data}</div>
                <div class="stat-label">With Grade Data</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">{exceeding_count}</div>
                <div class="stat-label">Exceeding Targets</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">{below_count}</div>
                <div class="stat-label">Below Targets</div>
            </div>
        </div>
"""
        
        # Priority students section
        if attention_needed:
            html_content += """
        <div class="section">
            <h2>🚨 Students Needing Priority Attention</h2>
            <div class="alert-attention">
                <strong>Action Required:</strong> These students have multiple subjects below their target grades and need immediate support.
            </div>
"""
            for student in attention_needed:
                priority_class = f"priority-{student['priority']}"
                html_content += f"""
            <div class="student-card {priority_class}">
                <div class="student-header">
                    <div class="student-name">👤 {student['name']}</div>
                    <div class="student-info">🏫 {student['school']} | 📅 {student['year']}</div>
                </div>
                <div class="performance-summary">
                    <strong>Performance Summary:</strong> 
                    🎉 Exceeding: {student['exceeding']} | 
                    ✅ Meeting: {student['meeting']} | 
                    ⚠️ Below: {student['below']}
                </div>
                <table class="grades-table">
                    <thead>
                        <tr>
                            <th>Subject</th>
                            <th>Current Grade</th>
                            <th>Target Grade</th>
                            <th>Status</th>
                        </tr>
                    </thead>
                    <tbody>
"""
                for subject_info in student['subjects']:
                    status_class = f"status-{subject_info['color'].replace('dark', '')}"
                    html_content += f"""
                        <tr>
                            <td>{subject_info['subject']}</td>
                            <td>{subject_info['current']}</td>
                            <td>{subject_info['predicted']}</td>
                            <td class="{status_class}">{subject_info['icon']} {subject_info['status']}</td>
                        </tr>
"""
                html_content += """
                    </tbody>
                </table>
            </div>
"""
            html_content += "</div>"
        
        # High performers section
        if high_performers:
            html_content += """
        <div class="section">
            <h2>🌟 High Performing Students</h2>
            <div class="alert-success">
                <strong>Excellent Work:</strong> These students are exceeding expectations and should be celebrated!
            </div>
"""
            for student in high_performers:
                html_content += f"""
            <div class="student-card priority-low">
                <div class="student-header">
                    <div class="student-name">⭐ {student['name']}</div>
                    <div class="student-info">🏫 {student['school']} | 📅 {student['year']}</div>
                </div>
                <div class="performance-summary">
                    <strong>Performance Summary:</strong> 
                    🎉 Exceeding: {student['exceeding']} | 
                    ✅ Meeting: {student['meeting']} | 
                    ⚠️ Below: {student['below']}
                </div>
                <table class="grades-table">
                    <thead>
                        <tr>
                            <th>Subject</th>
                            <th>Current Grade</th>
                            <th>Target Grade</th>
                            <th>Status</th>
                        </tr>
                    </thead>
                    <tbody>
"""
                for subject_info in student['subjects']:
                    if "Exceeding" in subject_info['status']:
                        status_class = f"status-{subject_info['color'].replace('dark', '')}"
                        html_content += f"""
                        <tr>
                            <td>{subject_info['subject']}</td>
                            <td>{subject_info['current']}</td>
                            <td>{subject_info['predicted']}</td>
                            <td class="{status_class}">{subject_info['icon']} {subject_info['status']}</td>
                        </tr>
"""
                html_content += """
                    </tbody>
                </table>
            </div>
"""
            html_content += "</div>"
        
        # All students section
        html_content += """
        <div class="section">
            <h2>📊 All Students Overview</h2>
"""
        
        for student in students_data:
            if not student['subjects']:
                continue
                
            priority_class = f"priority-{student['priority']}"
            html_content += f"""
            <div class="student-card {priority_class}">
                <div class="student-header">
                    <div class="student-name">👤 {student['name']}</div>
                    <div class="student-info">🏫 {student['school']} | 📅 {student['year']}</div>
                </div>
                <table class="grades-table">
                    <thead>
                        <tr>
                            <th>Subject</th>
                            <th>Current Grade</th>
                            <th>Target Grade</th>
                            <th>Status</th>
                        </tr>
                    </thead>
                    <tbody>
"""
            for subject_info in student['subjects']:
                status_class = f"status-{subject_info['color'].replace('dark', '')}"
                html_content += f"""
                        <tr>
                            <td>{subject_info['subject']}</td>
                            <td>{subject_info['current']}</td>
                            <td>{subject_info['predicted']}</td>
                            <td class="{status_class}">{subject_info['icon']} {subject_info['status']}</td>
                        </tr>
"""
            html_content += """
                    </tbody>
                </table>
            </div>
"""
        
        html_content += """
        </div>
        
        <div class="section">
            <h2>💡 Recommendations</h2>
            <div style="background-color: #f8f9fa; padding: 20px; border-radius: 8px;">
                <ol style="line-height: 1.8;">
                    <li><strong>🎯 Priority Interventions:</strong> Focus immediate support on students with multiple subjects below target</li>
                    <li><strong>📚 Subject-Specific Support:</strong> Provide additional tutoring in commonly struggling subjects (especially Maths)</li>
                    <li><strong>🏆 Celebrate Success:</strong> Recognize and reward high-performing students to maintain motivation</li>
                    <li><strong>📈 Regular Reviews:</strong> Implement weekly progress checks for Year 13 students approaching final exams</li>
                    <li><strong>🤝 Peer Support:</strong> Establish mentoring partnerships between high performers and struggling students</li>
                    <li><strong>👨‍👩‍👧‍👦 Parent Engagement:</strong> Schedule meetings with parents of students needing attention</li>
                </ol>
            </div>
        </div>
    </div>
</body>
</html>
"""
        
        # Save HTML file
        with open('Student_Grade_Analysis_Report.html', 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        print("✅ HTML Report generated: Student_Grade_Analysis_Report.html")
        return students_data
        
    except Exception as e:
        print(f"Error generating HTML report: {e}")
        import traceback
        traceback.print_exc()
        return []

def generate_excel_report(students_data):
    """Generate Excel report with multiple sheets"""
    try:
        with pd.ExcelWriter('Student_Grade_Analysis_Report.xlsx', engine='openpyxl') as writer:
            
            # Summary sheet
            summary_data = []
            for student in students_data:
                if student['subjects']:
                    summary_data.append({
                        'Student Name': student['name'],
                        'School': student['school'],
                        'Year': student['year'],
                        'Subjects Exceeding Target': student['exceeding'],
                        'Subjects Meeting Target': student['meeting'],
                        'Subjects Below Target': student['below'],
                        'Priority Level': student['priority'].title(),
                        'Total Subjects': len(student['subjects'])
                    })
            
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Summary', index=False)
            
            # Detailed grades sheet
            detailed_data = []
            for student in students_data:
                for subject_info in student['subjects']:
                    detailed_data.append({
                        'Student Name': student['name'],
                        'School': student['school'],
                        'Year': student['year'],
                        'Subject': subject_info['subject'],
                        'Current Grade': subject_info['current'],
                        'Target Grade': subject_info['predicted'],
                        'Status': subject_info['status'],
                        'Priority Level': student['priority'].title()
                    })
            
            detailed_df = pd.DataFrame(detailed_data)
            detailed_df.to_excel(writer, sheet_name='Detailed Grades', index=False)
            
            # Priority students sheet
            priority_data = []
            for student in students_data:
                if student['priority'] in ['high', 'medium'] and student['below'] > 0:
                    for subject_info in student['subjects']:
                        if 'Below' in subject_info['status']:
                            priority_data.append({
                                'Student Name': student['name'],
                                'School': student['school'],
                                'Year': student['year'],
                                'Subject': subject_info['subject'],
                                'Current Grade': subject_info['current'],
                                'Target Grade': subject_info['predicted'],
                                'Gap': f"{subject_info['current']} → {subject_info['predicted']}",
                                'Priority': student['priority'].title()
                            })
            
            if priority_data:
                priority_df = pd.DataFrame(priority_data)
                priority_df.to_excel(writer, sheet_name='Priority Students', index=False)
        
        print("✅ Excel Report generated: Student_Grade_Analysis_Report.xlsx")
        
    except Exception as e:
        print(f"Error generating Excel report: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    print("🎓 Generating Shareable Grade Analysis Reports...")
    print("=" * 60)
    
    # Generate HTML report
    students_data = generate_html_report()
    
    # Generate Excel report
    if students_data:
        generate_excel_report(students_data)
    
    print("\n📁 Files Generated:")
    print("   📄 Student_Grade_Analysis_Report.html - Interactive web report")
    print("   📊 Student_Grade_Analysis_Report.xlsx - Excel spreadsheet")
    print("\n💡 You can now share these files with your team!")
    print("   • Open the HTML file in any web browser")
    print("   • Open the Excel file in Microsoft Excel or Google Sheets")