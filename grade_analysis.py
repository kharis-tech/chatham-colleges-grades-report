import pandas as pd
import re
from typing import Dict, List, Tuple

def parse_grades(grade_text: str) -> Dict[str, str]:
    """Parse grade text and extract subject-grade pairs"""
    if pd.isna(grade_text) or grade_text == '-' or not grade_text.strip():
        return {}
    
    grades = {}
    # Split by newlines and process each line
    lines = str(grade_text).split('\n')
    
    for line in lines:
        line = line.strip()
        if not line or line == '-':
            continue
            
        # Look for patterns like "Subject - Grade" or "Subject: Grade" or "Subject Grade"
        # Handle various formats including A*, Merit, numbers, etc.
        patterns = [
            r'([^-:]+?)\s*[-:]\s*([A-Z*]+\d*|Merit|Distinction|\d+)',
            r'([^-:]+?)\s+([A-Z*]+\d*|Merit|Distinction|\d+)(?:\s|$)',
            r'([^(]+?)\s*\(([^)]+)\)',  # Subject (Grade) format
        ]
        
        matched = False
        for pattern in patterns:
            match = re.search(pattern, line, re.IGNORECASE)
            if match:
                subject = match.group(1).strip()
                grade = match.group(2).strip()
                grades[subject] = grade
                matched = True
                break
        
        # If no pattern matched, try to extract manually
        if not matched and line:
            # Look for common grade indicators
            grade_indicators = ['A*', 'A', 'B', 'C', 'D', 'E', 'U', 'Merit', 'Distinction', 'Pass']
            for indicator in grade_indicators:
                if indicator.lower() in line.lower():
                    subject = line.replace(indicator, '').strip(' -:')
                    if subject:
                        grades[subject] = indicator
                    break
    
    return grades

def analyze_grades():
    """Analyze grades from the Excel file"""
    try:
        # Read the grade tracker file
        df = pd.read_excel('KOC Grade Tracker Form(1-52).xlsx', sheet_name='Sheet1')
        
        print("=== UK STUDENT GRADE ANALYSIS ===\n")
        print("Comparing Current Grades vs Predicted/Aspiring Grades\n")
        print("=" * 80)
        
        for index, row in df.iterrows():
            name = row['Full Name']
            school = row['School You Attend']
            year = row['What year are you in']
            current_grades_text = row['Please list all the subjects you are currently taking and your current grades']
            predicted_grades_text = row['Please list all your predicted grades for each subject']
            
            if pd.isna(name) or name.strip() == '':
                continue
                
            print(f"\n📚 STUDENT: {name}")
            print(f"🏫 School: {school}")
            print(f"📅 Year: {year}")
            print("-" * 60)
            
            # Parse current and predicted grades
            current_grades = parse_grades(current_grades_text)
            predicted_grades = parse_grades(predicted_grades_text)
            
            if not current_grades and not predicted_grades:
                print("❌ No grade data available")
                continue
            
            # Get all subjects from both current and predicted
            all_subjects = set(current_grades.keys()) | set(predicted_grades.keys())
            
            if not all_subjects:
                print("❌ No subjects found in grade data")
                continue
            
            print(f"📊 GRADE COMPARISON:")
            print(f"{'Subject':<25} {'Current':<12} {'Predicted':<12} {'Status'}")
            print("-" * 60)
            
            for subject in sorted(all_subjects):
                current = current_grades.get(subject, 'N/A')
                predicted = predicted_grades.get(subject, 'N/A')
                
                # Determine status
                status = ""
                if current != 'N/A' and predicted != 'N/A':
                    if current == predicted:
                        status = "✅ On Track"
                    else:
                        # Simple comparison for UK grades (A* > A > B > C > D > E > U)
                        uk_grades = ['U', 'E', 'D', 'C', 'B', 'A', 'A*']
                        try:
                            current_val = uk_grades.index(current) if current in uk_grades else -1
                            predicted_val = uk_grades.index(predicted) if predicted in uk_grades else -1
                            
                            if current_val > predicted_val:
                                status = "🎉 Exceeding"
                            elif current_val < predicted_val:
                                status = "⚠️ Below Target"
                            else:
                                status = "✅ On Track"
                        except:
                            status = "📊 Compare"
                elif current != 'N/A':
                    status = "📈 Current Only"
                elif predicted != 'N/A':
                    status = "🎯 Target Set"
                
                print(f"{subject:<25} {current:<12} {predicted:<12} {status}")
            
            # Show raw data for debugging if needed
            print(f"\n📝 Raw Current Grades: {current_grades_text}")
            print(f"🎯 Raw Predicted Grades: {predicted_grades_text}")
            print("=" * 80)
        
    except Exception as e:
        print(f"Error analyzing grades: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    analyze_grades()