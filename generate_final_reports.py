import pandas as pd
import json
from datetime import datetime
from improved_standardization import ImprovedGradeStandardizer

def grade_to_points(grade: str) -> int:
    """Convert grades to points for comparison"""
    grade = grade.upper().strip()
    
    # GCSE 9-1 system
    gcse_points = {'9': 9, '8': 8, '7': 7, '6': 6, '5': 5, '4': 4, '3': 3, '2': 2, '1': 1, 'U': 0}
    
    # A-Level system
    alevel_points = {'A*': 6, 'A': 5, 'B': 4, 'C': 3, 'D': 2, 'E': 1, 'U': 0}
    
    # BTEC system
    btec_points = {'D*': 4, 'DISTINCTION': 4, 'MERIT': 2, 'PASS': 1}
    
    if grade in gcse_points:
        return gcse_points[grade]
    elif grade in alevel_points:
        return alevel_points[grade]
    elif grade in btec_points:
        return btec_points[grade]
    
    return -1

def compare_grades(current: str, predicted: str) -> tuple:
    """Compare current vs predicted grades"""
    if current == 'N/A' and predicted == 'N/A':
        return "No Data", "gray", "❓"
    elif current == 'N/A':
        return "Target Set", "blue", "🎯"
    elif predicted == 'N/A':
        return "Current Only", "green", "📈"
    
    current_points = grade_to_points(current)
    predicted_points = grade_to_points(predicted)
    
    if current_points == -1 or predicted_points == -1:
        if current.upper() == predicted.upper():
            return "Meeting Target", "green", "✅"
        else:
            return "Different Systems", "yellow", "📊"
    
    if current_points > predicted_points:
        return "Exceeding Target", "darkgreen", "🎉"
    elif current_points == predicted_points:
        return "Meeting Target", "green", "✅"
    else:
        return "Below Target", "red", "⚠️"

def generate_comprehensive_html_report(standardized_data):
    """Generate the final HTML report using standardized data"""
    
    # Process data for analysis
    students_analysis = []
    total_exceeding = 0
    total_meeting = 0
    total_below = 0
    all_subjects = set()
    
    for student in standardized_data:
        if not student['subjects']:
            continue
            
        analysis = {
            'name': student['name'],
            'school': student['school'],
            'year': student['year'],
            'subjects': [],
            'exceeding': 0,
            'meeting': 0,
            'below': 0,
            'priority': 'low'
        }
        
        for subject, grades in student['subjects'].items():
            current = grades['current']
            predicted = grades['predicted']
            status, color, icon = compare_grades(current, predicted)
            
            analysis['subjects'].append({
                'subject': subject,
                'current': current,
                'predicted': predicted,
                'status': status,
                'color': color,
                'icon': icon
            })
            
            all_subjects.add(subject)
            
            if "Exceeding" in status:
                analysis['exceeding'] += 1
                total_exceeding += 1
            elif "Meeting" in status:
                analysis['meeting'] += 1
                total_meeting += 1
            elif "Below" in status:
                analysis['below'] += 1
                total_below += 1
        
        # Determine priority
        if analysis['below'] >= 3:
            analysis['priority'] = 'high'
        elif analysis['below'] >= 1:
            analysis['priority'] = 'medium'
        
        students_analysis.append(analysis)
        students_analysis = sorted(students_analysis, key=lambda x: x["name"])
    
    # Generate HTML
    html_content = f"""
<!DOCTYPE html>
<html>
<head>
    <title>Colleges Grade Analysis Report</title>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 20px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
        }}
        .container {{
            max-width: 1200px;
            margin: 0 auto;
            background-color: white;
            padding: 30px;
            border-radius: 15px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.2);
        }}
        .header {{
            text-align: center;
            margin-bottom: 40px;
            border-bottom: 3px solid #2c3e50;
            padding-bottom: 20px;
        }}
        .header h1 {{
            color: #2c3e50;
            margin: 0;
            font-size: 2.8em;
            font-weight: 700;
        }}
        .header p {{
            color: #7f8c8d;
            margin: 15px 0 0 0;
            font-size: 1.2em;
        }}
        .summary-stats {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 25px;
            margin-bottom: 40px;
        }}
        .stat-card {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 25px;
            border-radius: 15px;
            text-align: center;
            box-shadow: 0 5px 15px rgba(0,0,0,0.1);
            transition: transform 0.3s ease;
        }}
        .stat-card:hover {{
            transform: translateY(-5px);
        }}
        .stat-number {{
            font-size: 3em;
            font-weight: bold;
            margin-bottom: 10px;
        }}
        .stat-label {{
            font-size: 1em;
            opacity: 0.9;
        }}
        .section {{
            margin-bottom: 50px;
        }}
        .section h2 {{
            color: #2c3e50;
            border-left: 6px solid #3498db;
            padding-left: 20px;
            margin-bottom: 25px;
            font-size: 1.8em;
        }}
        .student-card {{
            background-color: #f8f9fa;
            border: 1px solid #dee2e6;
            border-radius: 12px;
            padding: 25px;
            margin-bottom: 25px;
            transition: box-shadow 0.3s ease;
        }}
        .student-card:hover {{
            box-shadow: 0 5px 15px rgba(0,0,0,0.1);
        }}
        .student-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
            flex-wrap: wrap;
        }}
        .student-name {{
            font-size: 1.4em;
            font-weight: bold;
            color: #2c3e50;
        }}
        .student-info {{
            color: #6c757d;
            font-size: 1em;
        }}
        .grades-table {{
            width: 100%;
            border-collapse: collapse;
            margin-top: 15px;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }}
        .grades-table th, .grades-table td {{
            padding: 12px 15px;
            text-align: left;
            border-bottom: 1px solid #dee2e6;
        }}
        .grades-table th {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            font-weight: 600;
        }}
        .grades-table tr:nth-child(even) {{
            background-color: #f8f9fa;
        }}
        .grades-table tr:hover {{
            background-color: #e9ecef;
        }}
        .status-exceeding {{ color: #28a745; font-weight: bold; }}
        .status-meeting {{ color: #28a745; }}
        .status-below {{ color: #dc3545; font-weight: bold; }}
        .status-gray {{ color: #6c757d; }}
        .status-blue {{ color: #007bff; }}
        .status-green {{ color: #28a745; }}
        .status-yellow {{ color: #ffc107; }}
        .status-darkgreen {{ color: #155724; font-weight: bold; }}
        .status-red {{ color: #dc3545; font-weight: bold; }}
        
        .performance-summary {{
            background: linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%);
            border-left: 5px solid #2196f3;
            padding: 20px;
            margin-top: 20px;
            border-radius: 0 10px 10px 0;
        }}
        .alert-attention {{
            background: linear-gradient(135deg, #fff3cd 0%, #ffeaa7 100%);
            border: 1px solid #ffeaa7;
            border-radius: 10px;
            padding: 20px;
            margin-bottom: 25px;
        }}
        .alert-success {{
            background: linear-gradient(135deg, #d4edda 0%, #c3e6cb 100%);
            border: 1px solid #c3e6cb;
            border-radius: 10px;
            padding: 20px;
            margin-bottom: 25px;
        }}
        .priority-high {{ border-left: 6px solid #dc3545; }}
        .priority-medium {{ border-left: 6px solid #ffc107; }}
        .priority-low {{ border-left: 6px solid #28a745; }}
        
        .recommendations {{
            background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
            padding: 30px;
            border-radius: 15px;
            border: 1px solid #dee2e6;
        }}
        .recommendations ol {{
            line-height: 2;
            font-size: 1.1em;
        }}
        .recommendations li {{
            margin-bottom: 10px;
        }}
        
        @media (max-width: 768px) {{
            .container {{ padding: 15px; }}
            .student-header {{ flex-direction: column; align-items: flex-start; }}
            .grades-table {{ font-size: 0.9em; }}
            .stat-card {{ padding: 20px; }}
            .header h1 {{ font-size: 2.2em; }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🎓 Colleges Grade Analysis Report</h1>
            <p>Comprehensive Academic Performance Review</p>
            <p style="font-size: 0.9em; color: #95a5a6;">Generated on {datetime.now().strftime('%B %d, %Y at %I:%M %p')}</p>
        </div>
        
        <div class="summary-stats">
            <div class="stat-card">
                <div class="stat-number">{len(students_analysis)}</div>
                <div class="stat-label">Total Students</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">{len(all_subjects)}</div>
                <div class="stat-label">Unique Subjects</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">{total_exceeding}</div>
                <div class="stat-label">Exceeding Targets</div>
            </div>
            <div class="stat-card">
                <div class="stat-number">{total_below}</div>
                <div class="stat-label">Below Targets</div>
            </div>
        </div>
"""
    
    # All students overview
    html_content += """
        <div class="section">
            <h2>📊 Complete Student Overview</h2>
"""
    
    for student in students_analysis:
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
                            <td><strong>{subject_info['subject']}</strong></td>
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
    
    # Recommendations section
    html_content += f"""
        </div>
        
        <div class="section">
            <h2>💡 Strategic Recommendations</h2>
            <div class="recommendations">
                <ol>
                    <li><strong>📚 Subject-Specific Tutoring:</strong> Focus additional resources on commonly struggling subjects, particularly Mathematics and English</li>
                    <li><strong>📈 Progress Monitoring:</strong> Implement weekly progress reviews for Year 13 students approaching final examinations</li>
                    <li><strong>🤝 Peer Mentoring:</strong> Establish partnerships between high-achieving and struggling students</li>
                    <li><strong>👨‍👩‍👧‍👦 Parent Engagement:</strong> Schedule urgent meetings with parents of priority students</li>
                    <li><strong>📊 Data-Driven Decisions:</strong> Use this analysis to allocate teaching resources and plan intervention strategies</li>
                </ol>
            </div>
        </div>
        
        <div style="text-align: center; margin-top: 40px; padding: 20px; background-color: #f8f9fa; border-radius: 10px;">
            <p style="color: #6c757d; margin: 0;">
                <strong>Report Generated:</strong> {datetime.now().strftime('%B %d, %Y at %I:%M %p')} | 
                <strong>Total Students Analyzed:</strong> {len(students_analysis)} | 
                <strong>Subjects Tracked:</strong> {len(all_subjects)}
            </p>
        </div>
    </div>
</body>
</html>
"""
    
    # Save HTML file
    with open('Final_Student_Grade_Report.html', 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    return students_analysis

def generate_excel_report(students_analysis):
    """Generate comprehensive Excel report"""
    try:
        with pd.ExcelWriter('Final_Student_Grade_Report.xlsx', engine='openpyxl') as writer:
            
            # Summary sheet
            summary_data = []
            for student in students_analysis:
                summary_data.append({
                    'Student Name': student['name'],
                    'School': student['school'],
                    'Year': student['year'],
                    'Total Subjects': len(student['subjects']),
                    'Exceeding Target': student['exceeding'],
                    'Meeting Target': student['meeting'],
                    'Below Target': student['below'],
                    'Priority Level': student['priority'].title(),
                    'Needs Attention': 'Yes' if student['below'] >= 2 else 'No'
                })
            
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Student Summary', index=False)
            
            # Detailed grades sheet
            detailed_data = []
            for student in students_analysis:
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
            detailed_df.to_excel(writer, sheet_name='All Grades', index=False)
            
            # Priority students sheet
            priority_data = []
            for student in students_analysis:
                if student['below'] >= 2:
                    for subject_info in student['subjects']:
                        if 'Below' in subject_info['status']:
                            priority_data.append({
                                'Student Name': student['name'],
                                'School': student['school'],
                                'Year': student['year'],
                                'Subject': subject_info['subject'],
                                'Current Grade': subject_info['current'],
                                'Target Grade': subject_info['predicted'],
                                'Grade Gap': f"{subject_info['current']} → {subject_info['predicted']}",
                                'Action Required': 'High' if student['priority'] == 'high' else 'Medium'
                            })
            
            if priority_data:
                priority_df = pd.DataFrame(priority_data)
                priority_df.to_excel(writer, sheet_name='Priority Students', index=False)
        
        print("✅ Excel Report generated: Final_Student_Grade_Report.xlsx")
        
    except Exception as e:
        print(f"Error generating Excel report: {e}")

def main():
    """Main function to generate final reports"""
    print("🎓 GENERATING FINAL GRADE ANALYSIS REPORTS")
    print("=" * 70)
    
    try:
        # Load and standardize data
        df = pd.read_excel('KOC Grade Tracker Form(1-52).xlsx', sheet_name='Sheet1')
        standardizer = ImprovedGradeStandardizer()
        standardized_data = standardizer.process_all_data(df)
        
        print(f"✅ Standardized data for {len(standardized_data)} students")
        
        # Generate HTML report
        print("📄 Generating HTML report...")
        students_analysis = generate_comprehensive_html_report(standardized_data)
        print("✅ HTML Report generated: Final_Student_Grade_Report.html")
        
        # Generate Excel report
        print("📊 Generating Excel report...")
        generate_excel_report(students_analysis)
        
        print(f"\n🎉 REPORTS GENERATED SUCCESSFULLY!")
        print("=" * 70)
        print("📁 Files created:")
        print("   📄 Final_Student_Grade_Report.html - Beautiful web report")
        print("   📊 Final_Student_Grade_Report.xlsx - Comprehensive Excel analysis")
        print("\n💡 Share these files with your team - no technical knowledge required!")
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()