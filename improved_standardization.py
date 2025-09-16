import pandas as pd
import re
from typing import Dict, List, Tuple
import json

class ImprovedGradeStandardizer:
    def __init__(self):
        # Enhanced subject mappings with more variations
        self.subject_mappings = {
            # English variations
            'english lit': 'English Literature',
            'english literature': 'English Literature', 
            'english lang': 'English Language',
            'english language': 'English Language',
            'english': 'English Language',
            
            # Maths variations
            'maths': 'Mathematics',
            'mathematics': 'Mathematics',
            'math': 'Mathematics',
            
            # Sciences
            'biology': 'Biology',
            'chemistry': 'Chemistry', 
            'physics': 'Physics',
            'combined science': 'Combined Science',
            'applied science': 'Applied Science',
            'btec applied science': 'BTEC Applied Science',
            'science': 'Science',
            
            # Social subjects
            'history': 'History',
            'geography': 'Geography',
            'psychology': 'Psychology',
            'sociology': 'Sociology',
            'sociolgy': 'Sociology',  # Common typo
            'socio': 'Sociology',
            'religious studies': 'Religious Studies',
            'religious study': 'Religious Studies',
            're': 'Religious Studies',
            'ethics': 'Ethics',
            
            # Business & Economics
            'business': 'Business Studies',
            'business studies': 'Business Studies',
            'economics': 'Economics',
            'criminology': 'Criminology',
            'criminolgy': 'Criminology',  # Common typo
            'finance': 'Finance',
            
            # Languages
            'french': 'French',
            'spanish': 'Spanish', 
            'german': 'German',
            
            # Arts & Creative
            'art': 'Art',
            'music': 'Music',
            'drama': 'Drama',
            
            # Technology & Computing
            'ict': 'ICT',
            'it': 'ICT',
            'computing': 'Computing',
            'creative computing': 'Creative Computing',
            'computer science': 'Computer Science',
            
            # PE & Sports
            'pe': 'Physical Education',
            'physical education': 'Physical Education',
            'sport': 'Sport',
            'btec sport': 'BTEC Sport',
            'sports': 'Sport',
            
            # Health & Social Care
            'health and social care': 'Health & Social Care',
            'health & social care': 'Health & Social Care',
            'health and social': 'Health & Social Care',
            'child development': 'Child Development',
            'sports and nutrition': 'Sports & Nutrition',
            
            # Other subjects
            'politics': 'Politics',
            'media': 'Media Studies',
            'law': 'Law',
            'philosophy': 'Philosophy',
            'engineering': 'Engineering'
        }
    
    def extract_grades_robust(self, text: str) -> List[Tuple[str, str]]:
        """Robust grade extraction handling all formats"""
        if not text or pd.isna(text) or str(text).strip().lower() in ['nan', 'n/a', '-', 'na', '']:
            return []
        
        text = str(text).strip()
        pairs = []
        
        # Handle special single-grade cases first
        if re.match(r'^[A-Z*]{1,3}$', text.upper()):  # Like "AAA", "BBB", "A*"
            return [("Combined Subjects", text.upper())]
        
        if re.match(r'^\d$', text):  # Single number like "8"
            return [("General Target", text)]
        
        if text.lower() in ['merit', 'distinction', 'pass']:
            return [("General Grade", text.title())]
        
        # Split by newlines first, then by commas
        lines = []
        if '\n' in text:
            lines = [line.strip() for line in text.split('\n') if line.strip()]
        else:
            lines = [text]
        
        for line in lines:
            # Further split by commas if present
            if ',' in line and not re.search(r'[A-Za-z]+\s*,\s*[A-Za-z]+\s*-', line):
                parts = [part.strip() for part in line.split(',') if part.strip()]
            else:
                parts = [line]
            
            for part in parts:
                part_pairs = self._extract_from_part(part)
                pairs.extend(part_pairs)
        
        return pairs
    
    def _extract_from_part(self, part: str) -> List[Tuple[str, str]]:
        """Extract subject-grade pairs from a single part"""
        pairs = []
        part = part.strip()
        
        if not part:
            return pairs
        
        # Pattern 1: Subject - Grade (most common)
        pattern1 = r'([A-Za-z\s&\']+?)\s*[-–]\s*([A-Z*\d]+|Merit|Distinction|Pass|N/?A|NA|DMM|DDD|MMM|L2|Foundation)'
        matches = list(re.finditer(pattern1, part, re.IGNORECASE))
        
        if matches:
            for match in matches:
                subject = match.group(1).strip()
                grade = match.group(2).strip()
                if len(subject) > 1:  # Avoid single letters
                    pairs.append((subject, grade))
            return pairs
        
        # Pattern 2: Subject: Grade
        pattern2 = r'([A-Za-z\s&\']+?)\s*:\s*([A-Z*\d]+|Merit|Distinction|Pass|N/?A|NA|DMM|DDD|MMM|L2)'
        matches = list(re.finditer(pattern2, part, re.IGNORECASE))
        
        if matches:
            for match in matches:
                subject = match.group(1).strip()
                grade = match.group(2).strip()
                if len(subject) > 1:
                    pairs.append((subject, grade))
            return pairs
        
        # Pattern 3: Subject (Grade) format
        pattern3 = r'([A-Za-z\s&\']+?)\s*\(([A-Z*\d]+|Merit|Distinction|Pass)\)'
        matches = list(re.finditer(pattern3, part, re.IGNORECASE))
        
        if matches:
            for match in matches:
                subject = match.group(1).strip()
                grade = match.group(2).strip()
                if len(subject) > 1:
                    pairs.append((subject, grade))
            return pairs
        
        # Pattern 4: Try to find subjects and grades separately
        # Look for known subjects
        found_subject = None
        for subject_key in self.subject_mappings.keys():
            if subject_key in part.lower():
                found_subject = subject_key
                break
        
        if found_subject:
            # Look for grade in the remaining text
            remaining = part.lower().replace(found_subject, '').strip()
            grade_match = re.search(r'([A-Z*\d]+|Merit|Distinction|Pass|N/?A|NA)', remaining, re.IGNORECASE)
            if grade_match:
                pairs.append((found_subject, grade_match.group(1)))
            else:
                pairs.append((found_subject, "N/A"))
        
        # Pattern 5: If it contains grade-like patterns, treat as subject
        elif re.search(r'[A-Z*\d]+|Merit|Distinction|Pass', part, re.IGNORECASE):
            grade_match = re.search(r'([A-Z*\d]+|Merit|Distinction|Pass)', part, re.IGNORECASE)
            if grade_match:
                subject_part = part.replace(grade_match.group(1), '').strip(' -:')
                if subject_part:
                    pairs.append((subject_part, grade_match.group(1)))
        
        return pairs
    
    def standardize_subject(self, subject: str) -> str:
        """Standardize subject name"""
        if not subject:
            return "Unknown Subject"
        
        subject = subject.strip().lower()
        
        # Remove common prefixes/suffixes
        subject = re.sub(r'^(btec|level \d+|l\d+)\s*', '', subject)
        subject = re.sub(r'\s*(gcse|a-level|as)$', '', subject)
        subject = re.sub(r'\s*\([^)]*\)$', '', subject)
        subject = re.sub(r'[^\w\s&]', '', subject)  # Remove special chars except &
        subject = re.sub(r'\s+', ' ', subject).strip()
        
        # Direct mapping
        if subject in self.subject_mappings:
            return self.subject_mappings[subject]
        
        # Fuzzy matching
        for key, value in self.subject_mappings.items():
            if key in subject or subject in key:
                return value
        
        # Title case if no match
        return subject.title() if subject else "Unknown Subject"
    
    def standardize_grade(self, grade: str) -> str:
        """Standardize grade"""
        if not grade or pd.isna(grade):
            return "N/A"
        
        grade = str(grade).strip()
        
        # Handle N/A cases
        if grade.lower() in ['na', 'n/a', '-', 'nan', 'no grade', 'not available']:
            return "N/A"
        
        # BTEC grades
        if grade.lower() in ['merit', 'm']:
            return "Merit"
        if grade.lower() in ['distinction', 'd']:
            return "Distinction"
        if grade.lower() in ['pass', 'p']:
            return "Pass"
        if grade.lower() in ['distinction*', 'd*']:
            return "D*"
        
        # Combined BTEC grades
        if grade.upper() in ['DMM', 'DDD', 'MMM', 'PPP', 'DD*', 'DM', 'MP']:
            return grade.upper()
        
        # A-Level grades
        if grade.upper() in ['A*', 'A', 'B', 'C', 'D', 'E', 'U']:
            return grade.upper()
        
        # GCSE grades
        if grade in ['9', '8', '7', '6', '5', '4', '3', '2', '1']:
            return grade
        
        # Level 2
        if grade.upper() in ['L2', 'LEVEL 2']:
            return "L2"
        
        return grade.upper()
    
    def process_all_data(self, df: pd.DataFrame) -> List[Dict]:
        """Process all student data"""
        standardized_data = []
        
        print(f"Processing {len(df)} rows...")
        
        for index, row in df.iterrows():
            name = row['Full Name']
            if pd.isna(name) or not str(name).strip():
                continue
            
            student = {
                'name': str(name).strip(),
                'school': str(row['School You Attend']).strip() if not pd.isna(row['School You Attend']) else "Unknown",
                'year': str(row['What year are you in']).strip() if not pd.isna(row['What year are you in']) else "Unknown",
                'subjects': {},
                'raw_current': str(row['Please list all the subjects you are currently taking and your current grades']) if not pd.isna(row['Please list all the subjects you are currently taking and your current grades']) else "",
                'raw_predicted': str(row['Please list all your predicted grades for each subject']) if not pd.isna(row['Please list all your predicted grades for each subject']) else ""
            }
            
            # Process current grades
            current_pairs = self.extract_grades_robust(student['raw_current'])
            for subject, grade in current_pairs:
                std_subject = self.standardize_subject(subject)
                std_grade = self.standardize_grade(grade)
                
                if std_subject not in student['subjects']:
                    student['subjects'][std_subject] = {'current': 'N/A', 'predicted': 'N/A'}
                student['subjects'][std_subject]['current'] = std_grade
            
            # Process predicted grades
            predicted_pairs = self.extract_grades_robust(student['raw_predicted'])
            for subject, grade in predicted_pairs:
                std_subject = self.standardize_subject(subject)
                std_grade = self.standardize_grade(grade)
                
                if std_subject not in student['subjects']:
                    student['subjects'][std_subject] = {'current': 'N/A', 'predicted': 'N/A'}
                student['subjects'][std_subject]['predicted'] = std_grade
            
            standardized_data.append(student)
        
        return standardized_data

def main():
    """Main standardization function"""
    try:
        print("🔧 IMPROVED GRADE DATA STANDARDIZATION")
        print("=" * 60)
        
        # Load data
        df = pd.read_excel('KOC Grade Tracker Form(1-52).xlsx', sheet_name='Sheet1')
        print(f"Loaded {len(df)} rows from Excel file")
        
        # Initialize standardizer
        standardizer = ImprovedGradeStandardizer()
        
        # Process all data
        standardized_data = standardizer.process_all_data(df)
        
        # Save to JSON
        with open('standardized_grades.json', 'w', encoding='utf-8') as f:
            json.dump(standardized_data, f, indent=2, ensure_ascii=False)
        
        print(f"\n✅ Processed {len(standardized_data)} students")
        print("✅ Saved to: standardized_grades.json")
        
        # Show summary
        total_subjects = set()
        students_with_data = 0
        
        print(f"\n📊 SAMPLE RESULTS (First 3 students):")
        print("-" * 50)
        
        for i, student in enumerate(standardized_data[:3]):
            if student['subjects']:
                students_with_data += 1
                print(f"\n{i+1}. {student['name']} ({student['year']})")
                for subject, grades in student['subjects'].items():
                    print(f"   • {subject}: {grades['current']} → {grades['predicted']}")
                    total_subjects.add(subject)
        
        print(f"\n📈 SUMMARY:")
        print(f"   Students with grade data: {students_with_data}")
        print(f"   Unique subjects found: {len(total_subjects)}")
        
        return standardized_data
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return []

if __name__ == "__main__":
    main()