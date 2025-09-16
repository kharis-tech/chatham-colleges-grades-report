import pandas as pd
import re
from typing import Dict, List, Tuple
import json

class GradeDataStandardizer:
    def __init__(self):
        # UK subject name standardization mapping
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
            
            # Social subjects
            'history': 'History',
            'geography': 'Geography',
            'psychology': 'Psychology',
            'sociology': 'Sociology',
            'sociolgy': 'Sociology',  # Common typo
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
            
            # Health & Social Care
            'health and social care': 'Health & Social Care',
            'health & social care': 'Health & Social Care',
            'health and social': 'Health & Social Care',
            'child development': 'Child Development',
            
            # Other subjects
            'politics': 'Politics',
            'media': 'Media Studies',
            'law': 'Law',
            'philosophy': 'Philosophy',
            'engineering': 'Engineering',
            'finance': 'Finance'
        }
        
        # Grade standardization
        self.grade_mappings = {
            # GCSE grades
            '9': '9', '8': '8', '7': '7', '6': '6', '5': '5', '4': '4', '3': '3', '2': '2', '1': '1',
            
            # A-Level grades
            'a*': 'A*', 'a': 'A', 'b': 'B', 'c': 'C', 'd': 'D', 'e': 'E', 'u': 'U',
            
            # BTEC grades
            'distinction*': 'D*', 'd*': 'D*', 'distinction': 'D', 'd': 'D', 
            'merit': 'M', 'm': 'M', 'pass': 'P', 'p': 'P',
            
            # Special cases
            'dmm': 'DMM', 'ddd': 'DDD', 'mmm': 'MMM', 'ppp': 'PPP',
            'dd*': 'DD*', 'dm': 'DM', 'mp': 'MP',
            
            # Level 2 qualifications
            'l2': 'L2', 'level 2': 'L2',
            
            # Not available
            'na': 'N/A', 'n/a': 'N/A', 'nan': 'N/A', '-': 'N/A', '': 'N/A'
        }
    
    def clean_text(self, text: str) -> str:
        """Clean and normalize text"""
        if pd.isna(text) or not text:
            return ""
        
        text = str(text).strip()
        
        # Remove common prefixes that don't add value
        text = re.sub(r'^(please list all|current grades?:?|predicted grades?:?)', '', text, flags=re.IGNORECASE)
        text = re.sub(r'\s+', ' ', text)  # Normalize whitespace
        
        return text.strip()
    
    def extract_subject_grade_pairs(self, text: str) -> List[Tuple[str, str]]:
        """Extract subject-grade pairs from text using multiple strategies"""
        if not text or text.lower() in ['nan', 'n/a', '-', 'na']:
            return []
        
        pairs = []
        
        # Strategy 1: Handle newline-separated entries
        lines = text.split('\n')
        if len(lines) > 1:
            for line in lines:
                line = line.strip()
                if line:
                    line_pairs = self._parse_single_line(line)
                    pairs.extend(line_pairs)
        else:
            # Strategy 2: Handle comma-separated entries
            if ',' in text:
                parts = text.split(',')
                for part in parts:
                    part = part.strip()
                    if part:
                        part_pairs = self._parse_single_line(part)
                        pairs.extend(part_pairs)
            else:
                # Strategy 3: Single line parsing
                pairs = self._parse_single_line(text)
        
        return pairs
    
    def _parse_single_line(self, line: str) -> List[Tuple[str, str]]:
        """Parse a single line for subject-grade pairs"""
        pairs = []
        line = line.strip()
        
        if not line:
            return pairs
        
        # Pattern 1: Subject - Grade (most common)
        pattern1 = r'([A-Za-z\s&]+?)\s*[-–]\s*([A-Z*\d]+|Merit|Distinction|Pass|N/?A|NA|DMM|DDD|MMM|L2)'
        matches = re.finditer(pattern1, line, re.IGNORECASE)
        for match in matches:
            subject = match.group(1).strip()
            grade = match.group(2).strip()
            if subject and grade:
                pairs.append((subject, grade))
        
        # If no matches with pattern 1, try other patterns
        if not pairs:
            # Pattern 2: Subject: Grade
            pattern2 = r'([A-Za-z\s&]+?)\s*:\s*([A-Z*\d]+|Merit|Distinction|Pass|N/?A|NA|DMM|DDD|MMM|L2)'
            matches = re.finditer(pattern2, line, re.IGNORECASE)
            for match in matches:
                subject = match.group(1).strip()
                grade = match.group(2).strip()
                if subject and grade:
                    pairs.append((subject, grade))
        
        # If still no matches, try to extract from context
        if not pairs:
            # Pattern 3: Just grades (like "AAA" or "BBB")
            if re.match(r'^[A-Z*]+$', line.upper()) and len(line) <= 5:
                pairs.append(("Combined Subjects", line.upper()))
            
            # Pattern 4: Just a number (like "8" for predicted grade)
            elif re.match(r'^\d$', line):
                pairs.append(("General Target", line))
            
            # Pattern 5: Subject without explicit grade
            elif any(subj in line.lower() for subj in self.subject_mappings.keys()):
                # Try to find a grade in the line
                grade_match = re.search(r'([A-Z*\d]+|Merit|Distinction|Pass)', line, re.IGNORECASE)
                if grade_match:
                    pairs.append((line.replace(grade_match.group(1), '').strip(), grade_match.group(1)))
                else:
                    pairs.append((line, "N/A"))
        
        return pairs
    
    def standardize_subject_name(self, subject: str) -> str:
        """Standardize subject names"""
        if not subject:
            return "Unknown Subject"
        
        subject = subject.strip().lower()
        
        # Remove common prefixes/suffixes
        subject = re.sub(r'^(btec|level \d+|l\d+)\s*', '', subject)
        subject = re.sub(r'\s*(gcse|a-level|as)$', '', subject)
        subject = re.sub(r'\s*\([^)]*\)$', '', subject)  # Remove parentheses content
        
        # Direct mapping
        if subject in self.subject_mappings:
            return self.subject_mappings[subject]
        
        # Fuzzy matching for common variations
        for key, value in self.subject_mappings.items():
            if key in subject or subject in key:
                return value
        
        # If no match found, title case the original
        return subject.title()
    
    def standardize_grade(self, grade: str) -> str:
        """Standardize grade format"""
        if not grade:
            return "N/A"
        
        grade = str(grade).strip().lower()
        
        # Direct mapping
        if grade in self.grade_mappings:
            return self.grade_mappings[grade]
        
        # Handle special cases
        if grade in ['no grade', 'no current grade', 'not available']:
            return "N/A"
        
        # If it's already in correct format, return as is (but uppercase)
        if re.match(r'^[A-Z*\d]+$', grade.upper()):
            return grade.upper()
        
        return grade.upper()
    
    def process_student_data(self, df: pd.DataFrame) -> List[Dict]:
        """Process all student data and return standardized format"""
        standardized_data = []
        
        for index, row in df.iterrows():
            name = row['Full Name']
            if pd.isna(name) or name.strip() == '':
                continue
            
            student_data = {
                'name': name.strip(),
                'school': str(row['School You Attend']).strip() if not pd.isna(row['School You Attend']) else "Unknown School",
                'year': str(row['What year are you in']).strip() if not pd.isna(row['What year are you in']) else "Unknown Year",
                'subjects': {}
            }
            
            # Process current grades
            current_text = self.clean_text(row['Please list all the subjects you are currently taking and your current grades'])
            current_pairs = self.extract_subject_grade_pairs(current_text)
            
            # Process predicted grades
            predicted_text = self.clean_text(row['Please list all your predicted grades for each subject'])
            predicted_pairs = self.extract_subject_grade_pairs(predicted_text)
            
            # Combine all subjects
            all_subjects = set()
            
            # Add current grades
            for subject, grade in current_pairs:
                std_subject = self.standardize_subject_name(subject)
                std_grade = self.standardize_grade(grade)
                all_subjects.add(std_subject)
                
                if std_subject not in student_data['subjects']:
                    student_data['subjects'][std_subject] = {'current': 'N/A', 'predicted': 'N/A'}
                student_data['subjects'][std_subject]['current'] = std_grade
            
            # Add predicted grades
            for subject, grade in predicted_pairs:
                std_subject = self.standardize_subject_name(subject)
                std_grade = self.standardize_grade(grade)
                all_subjects.add(std_subject)
                
                if std_subject not in student_data['subjects']:
                    student_data['subjects'][std_subject] = {'current': 'N/A', 'predicted': 'N/A'}
                student_data['subjects'][std_subject]['predicted'] = std_grade
            
            # Store raw data for debugging
            student_data['raw_current'] = current_text
            student_data['raw_predicted'] = predicted_text
            
            standardized_data.append(student_data)
        
        return standardized_data
    
    def save_standardized_data(self, standardized_data: List[Dict], filename: str = 'standardized_grade_data.json'):
        """Save standardized data to JSON file"""
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(standardized_data, f, indent=2, ensure_ascii=False)
        print(f"✅ Standardized data saved to: {filename}")
    
    def create_standardization_report(self, standardized_data: List[Dict]):
        """Create a report showing the standardization results"""
        print("🔧 GRADE DATA STANDARDIZATION REPORT")
        print("=" * 80)
        
        total_students = len(standardized_data)
        total_subjects = set()
        total_current_grades = 0
        total_predicted_grades = 0
        
        print(f"\n📊 SUMMARY:")
        print(f"   Total Students Processed: {total_students}")
        
        # Show first few examples
        print(f"\n📝 STANDARDIZATION EXAMPLES (First 5 students):")
        print("-" * 60)
        
        for i, student in enumerate(standardized_data[:5]):
            print(f"\n{i+1}. {student['name']} ({student['school']}, {student['year']})")
            
            if student['subjects']:
                print("   Subjects:")
                for subject, grades in student['subjects'].items():
                    current = grades['current']
                    predicted = grades['predicted']
                    print(f"      • {subject}: {current} → {predicted}")
                    total_subjects.add(subject)
                    if current != 'N/A':
                        total_current_grades += 1
                    if predicted != 'N/A':
                        total_predicted_grades += 1
            else:
                print("   No grade data available")
            
            # Show raw data for comparison
            if student['raw_current']:
                print(f"   Raw Current: '{student['raw_current'][:100]}{'...' if len(student['raw_current']) > 100 else ''}'")
            if student['raw_predicted']:
                print(f"   Raw Predicted: '{student['raw_predicted'][:100]}{'...' if len(student['raw_predicted']) > 100 else ''}'")
        
        print(f"\n📈 STATISTICS:")
        print(f"   Unique Subjects Found: {len(total_subjects)}")
        print(f"   Total Current Grades: {total_current_grades}")
        print(f"   Total Predicted Grades: {total_predicted_grades}")
        
        print(f"\n📚 ALL STANDARDIZED SUBJECTS:")
        for subject in sorted(total_subjects):
            print(f"   • {subject}")

def main():
    """Main function to run standardization"""
    try:
        # Load data
        df = pd.read_excel('KOC Grade Tracker Form(1-52).xlsx', sheet_name='Sheet1')
        
        # Initialize standardizer
        standardizer = GradeDataStandardizer()
        
        # Process data
        print("🔧 Starting grade data standardization...")
        standardized_data = standardizer.process_student_data(df)
        
        # Save standardized data
        standardizer.save_standardized_data(standardized_data)
        
        # Create report
        standardizer.create_standardization_report(standardized_data)
        
        return standardized_data
        
    except Exception as e:
        print(f"Error in standardization: {e}")
        import traceback
        traceback.print_exc()
        return []

if __name__ == "__main__":
    main()