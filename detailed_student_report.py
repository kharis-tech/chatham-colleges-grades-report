import pandas as pd
import re
from typing import Dict, List, Tuple

def parse_subject_grades(text: str) -> Dict[str, str]:
    """Parse subject-grade pairs from text"""
    if pd.isna(text) or not str(text).strip() or str(text).lower() in ['nan', '-', 'n/a']:
        return {}
    
    grades = {}
    text = str(text).strip()
    
    # Common patterns for UK grades
    patterns = [
        # Subject - Grade
        r'([A-Za-z\s&]+?)\s*[-–]\s*([A-Z*]+\d*|Merit|Distinction|Pass|\d+)',
        # Subject: Grade  
        r'([A-Za-z\s&]+?)\s*:\s*([A-Z*]+\d*|Merit|Distinction|Pass|\d+)',
        # Subject Grade (space)
        r'([A-Za-z\s&]+?)\s+([A-Z*]+\d*|Merit|Distinction|Pass|\d+)(?=\s|$|,)',
    ]
    
    lines = re.split(r'[,\n]', text)
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        matched = False
        for pattern in patterns:
            matches = re.finditer(pattern, line, re.IGNORECASE)
            for match in matches:
                subject = match.group(1).strip()
                grade = match.group(2).strip()
                
                # Clean subject name
                subject = re.sub(r'\s+', ' ', subject)
                subject = subject.title()
                
                if len(subject) > 1 and grade:
                    grades[subject] = grade
                    matched = True
        
        # If no pattern matched but line contains subject keywords
        if not matched:
            subject_keywords = ['english', 'maths', 'science', 'biology', 'chemistry', 'physics', 
                              'history', 'geography', 'psychology', 'sociology', 'business', 
                              'economics', 'art', 'music', 'pe', 'sport', 'computing', 'it']
            
            for keyword in subject_keywords:
                if keyword in line.lower():
                    # Try to extract grade from the line
                    grade_match = re.search(r'([A-Z*]+\d*|Merit|Distinction|Pass|\d+)', line, re.IGNORECASE)
                    if grade_match:
                        grades[line.strip().title()] = grade_match.group(1)
                    break
    
    return grades

def grade_to_points(grade: str) -> int:
    """Convert UK grades to point system for comparison"""
    grade = grade.upper().strip()
    
    # GCSE 9-1 system
    gcse_points = {'9': 9, '8': 8, '7': 7, '6': 6, '5': 5, '4': 4, '3': 3, '2': 2, '1': 1, 'U': 0}
    
    # A-Level system  
    alevel_points = {'A*': 6, 'A': 5, 'B': 4, 'C': 3, 'D': 2, 'E': 1, 'U': 0}
    
    # BTEC system
    btec_points = {'D*': 4, 'DISTINCTION*': 4, 'D': 3, 'DISTINCTION': 3, 'M': 2, 'MERIT': 2, 'P': 1, 'PASS': 1}
    
    if grade in gcse_points:
        return gcse_points[grade]
    elif grade in alevel_points:
        return alevel_points[grade]
    elif grade in btec_points:
        return btec_points[grade]
    
    return -1  # Unknown grade

def compare_performance(current: str, predicted: str) -> Tuple[str, str]:
    """Compare current vs predicted performance"""
    if current == 'N/A' and predicted == 'N/A':
        return "❓ No Data", "gray"
    elif current == 'N/A':
        return "🎯 Target Set", "blue"
    elif predicted == 'N/A':
        return "📈 Has Current Grade", "green"
    
    current_points = grade_to_points(current)
    predicted_points = grade_to_points(predicted)
    
    if current_points == -1 or predicted_points == -1:
        if current.upper() == predicted.upper():
            return "✅ Matching", "green"
        else:
            return "📊 Different Format", "yellow"
    
    if current_points > predicted_points:
        return "🎉 Exceeding Target", "green"
    elif current_points == predicted_points:
        return "✅ Meeting Target", "green"
    else:
        return "⚠️ Below Target", "red"

def create_detailed_report():
    """Create detailed individual student reports"""
    try:
        df = pd.read_excel('KOC Grade Tracker Form(1-52).xlsx', sheet_name='Sheet1')
        
        print("🎓 DETAILED UK STUDENT GRADE ANALYSIS")
        print("=" * 100)
        
        for index, row in df.iterrows():
            name = row['Full Name']
            if pd.isna(name) or name.strip() == '':
                continue
            
            school = row['School You Attend']
            year = row['What year are you in']
            current_text = row['Please list all the subjects you are currently taking and your current grades']
            predicted_text = row['Please list all your predicted grades for each subject']
            
            print(f"\n👤 STUDENT: {name}")
            print(f"🏫 School: {school}")
            print(f"📅 Year: {year}")
            print("-" * 90)
            
            current_grades = parse_subject_grades(current_text)
            predicted_grades = parse_subject_grades(predicted_text)
            
            # Get all unique subjects
            all_subjects = set()
            for subject in current_grades.keys():
                all_subjects.add(subject)
            for subject in predicted_grades.keys():
                all_subjects.add(subject)
            
            if not all_subjects:
                print("❌ No grade data available")
                continue
            
            print(f"📊 GRADE COMPARISON:")
            print(f"{'Subject':<30} {'Current':<12} {'Predicted':<12} {'Status':<20}")
            print("-" * 90)
            
            performance_summary = {'exceeding': 0, 'meeting': 0, 'below': 0, 'no_data': 0}
            
            for subject in sorted(all_subjects):
                current = current_grades.get(subject, 'N/A')
                predicted = predicted_grades.get(subject, 'N/A')
                status, color = compare_performance(current, predicted)
                
                print(f"{subject:<30} {current:<12} {predicted:<12} {status:<20}")
                
                if "Exceeding" in status:
                    performance_summary['exceeding'] += 1
                elif "Meeting" in status or "Matching" in status:
                    performance_summary['meeting'] += 1
                elif "Below" in status:
                    performance_summary['below'] += 1
                else:
                    performance_summary['no_data'] += 1
            
            # Performance summary
            total_comparable = performance_summary['exceeding'] + performance_summary['meeting'] + performance_summary['below']
            if total_comparable > 0:
                print(f"\n📈 PERFORMANCE SUMMARY:")
                print(f"   🎉 Exceeding Targets: {performance_summary['exceeding']}")
                print(f"   ✅ Meeting Targets: {performance_summary['meeting']}")
                print(f"   ⚠️ Below Targets: {performance_summary['below']}")
                
                if performance_summary['below'] > 0:
                    print(f"   🚨 ATTENTION NEEDED: {performance_summary['below']} subjects below target")
                elif performance_summary['exceeding'] > 0:
                    print(f"   🌟 EXCELLENT: Exceeding targets in {performance_summary['exceeding']} subjects")
                else:
                    print(f"   👍 GOOD: Meeting all targets")
            
            print("=" * 100)
        
    except Exception as e:
        print(f"Error creating detailed report: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    create_detailed_report()