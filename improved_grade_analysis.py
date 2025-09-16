import pandas as pd
import re
from typing import Dict, List, Tuple

def clean_subject_name(subject: str) -> str:
    """Clean and standardize subject names"""
    subject = subject.strip()
    # Remove common prefixes/suffixes
    subject = re.sub(r'^(BTEC|Level \d+|L\d+)\s*', '', subject, flags=re.IGNORECASE)
    subject = re.sub(r'\s*(GCSE|A-Level|AS)$', '', subject, flags=re.IGNORECASE)
    return subject.title()

def parse_grades_improved(grade_text: str) -> Dict[str, str]:
    """Improved grade parsing for UK education system"""
    if pd.isna(grade_text) or grade_text == '-' or not grade_text.strip():
        return {}
    
    grades = {}
    text = str(grade_text).strip()
    
    # Handle special cases first
    if 'no grade' in text.lower() or 'n/a' in text.lower() or text.lower() == 'nan':
        return {}
    
    # Split by common delimiters
    lines = re.split(r'[,\n]', text)
    
    for line in lines:
        line = line.strip()
        if not line or line == '-':
            continue
        
        # Pattern 1: Subject - Grade (most common)
        match = re.search(r'([^-:]+?)\s*[-:]\s*([A-Z*]+\d*|Merit|Distinction|Pass|\d+)', line, re.IGNORECASE)
        if match:
            subject = clean_subject_name(match.group(1))
            grade = match.group(2).strip()
            grades[subject] = grade
            continue
        
        # Pattern 2: Subject Grade (space separated)
        match = re.search(r'([A-Za-z\s&]+?)\s+([A-Z*]+\d*|Merit|Distinction|Pass|\d+)(?:\s|$)', line, re.IGNORECASE)
        if match:
            subject = clean_subject_name(match.group(1))
            grade = match.group(2).strip()
            grades[subject] = grade
            continue
        
        # Pattern 3: Just subject names (when grades are mentioned separately)
        if any(keyword in line.lower() for keyword in ['english', 'maths', 'science', 'biology', 'chemistry', 'physics', 'history', 'geography', 'psychology', 'sociology', 'business', 'economics']):
            subject = clean_subject_name(line)
            if subject and len(subject) > 2:
                grades[subject] = 'Listed'
    
    return grades

def compare_grades(current: str, predicted: str) -> str:
    """Compare UK grades and return status"""
    if current == 'N/A' and predicted == 'N/A':
        return "❓ No Data"
    elif current == 'N/A':
        return "🎯 Target Set"
    elif predicted == 'N/A':
        return "📈 Current Only"
    
    # UK GCSE grades: 9 > 8 > 7 > 6 > 5 > 4 > 3 > 2 > 1 > U
    gcse_grades = {'U': 0, '1': 1, '2': 2, '3': 3, '4': 4, '5': 5, '6': 6, '7': 7, '8': 8, '9': 9}
    
    # UK A-Level grades: A* > A > B > C > D > E > U
    alevel_grades = {'U': 0, 'E': 1, 'D': 2, 'C': 3, 'B': 4, 'A': 5, 'A*': 6}
    
    # BTEC grades: Distinction* > Distinction > Merit > Pass
    btec_grades = {'Pass': 1, 'P': 1, 'Merit': 2, 'M': 2, 'Distinction': 3, 'D': 3, 'D*': 4, 'Distinction*': 4}
    
    current_val = None
    predicted_val = None
    
    # Try to match grades in different systems
    for grade_system in [gcse_grades, alevel_grades, btec_grades]:
        if current in grade_system and predicted in grade_system:
            current_val = grade_system[current]
            predicted_val = grade_system[predicted]
            break
    
    if current_val is not None and predicted_val is not None:
        if current_val > predicted_val:
            return "🎉 Exceeding Target"
        elif current_val == predicted_val:
            return "✅ On Track"
        else:
            return "⚠️ Below Target"
    
    # If we can't compare numerically, check if they're the same
    if current.upper() == predicted.upper():
        return "✅ On Track"
    
    return "📊 Different Systems"

def create_summary_report():
    """Create a comprehensive summary report"""
    try:
        df = pd.read_excel('KOC Grade Tracker Form(1-52).xlsx', sheet_name='Sheet1')
        
        print("🎓 UK STUDENT GRADE ANALYSIS SUMMARY")
        print("=" * 80)
        
        total_students = 0
        students_with_data = 0
        exceeding_count = 0
        on_track_count = 0
        below_target_count = 0
        
        detailed_results = []
        
        for index, row in df.iterrows():
            name = row['Full Name']
            if pd.isna(name) or name.strip() == '':
                continue
                
            total_students += 1
            
            school = row['School You Attend']
            year = row['What year are you in']
            current_grades_text = row['Please list all the subjects you are currently taking and your current grades']
            predicted_grades_text = row['Please list all your predicted grades for each subject']
            
            current_grades = parse_grades_improved(current_grades_text)
            predicted_grades = parse_grades_improved(predicted_grades_text)
            
            if current_grades or predicted_grades:
                students_with_data += 1
            
            # Get all subjects
            all_subjects = set(current_grades.keys()) | set(predicted_grades.keys())
            
            student_summary = {
                'name': name,
                'school': school,
                'year': year,
                'subjects': [],
                'exceeding': 0,
                'on_track': 0,
                'below_target': 0
            }
            
            for subject in sorted(all_subjects):
                current = current_grades.get(subject, 'N/A')
                predicted = predicted_grades.get(subject, 'N/A')
                status = compare_grades(current, predicted)
                
                student_summary['subjects'].append({
                    'subject': subject,
                    'current': current,
                    'predicted': predicted,
                    'status': status
                })
                
                if "Exceeding" in status:
                    student_summary['exceeding'] += 1
                    exceeding_count += 1
                elif "On Track" in status:
                    student_summary['on_track'] += 1
                    on_track_count += 1
                elif "Below Target" in status:
                    student_summary['below_target'] += 1
                    below_target_count += 1
            
            detailed_results.append(student_summary)
        
        # Print summary statistics
        print(f"\n📊 OVERALL STATISTICS")
        print(f"Total Students: {total_students}")
        print(f"Students with Grade Data: {students_with_data}")
        print(f"🎉 Subjects Exceeding Target: {exceeding_count}")
        print(f"✅ Subjects On Track: {on_track_count}")
        print(f"⚠️ Subjects Below Target: {below_target_count}")
        
        # Print detailed results for students with concerning performance
        print(f"\n⚠️ STUDENTS NEEDING ATTENTION (Below Target in Multiple Subjects)")
        print("=" * 80)
        
        for student in detailed_results:
            if student['below_target'] >= 2:  # Students with 2+ subjects below target
                print(f"\n👤 {student['name']} ({student['school']}, {student['year']})")
                print(f"   Below Target: {student['below_target']} subjects")
                for subject_info in student['subjects']:
                    if "Below Target" in subject_info['status']:
                        print(f"   • {subject_info['subject']}: {subject_info['current']} → {subject_info['predicted']} {subject_info['status']}")
        
        # Print high performers
        print(f"\n🎉 HIGH PERFORMERS (Exceeding Targets)")
        print("=" * 80)
        
        for student in detailed_results:
            if student['exceeding'] >= 1:  # Students exceeding in at least 1 subject
                print(f"\n⭐ {student['name']} ({student['school']}, {student['year']})")
                print(f"   Exceeding: {student['exceeding']} subjects")
                for subject_info in student['subjects']:
                    if "Exceeding" in subject_info['status']:
                        print(f"   • {subject_info['subject']}: {subject_info['current']} → {subject_info['predicted']} {subject_info['status']}")
        
        return detailed_results
        
    except Exception as e:
        print(f"Error creating summary: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    create_summary_report()