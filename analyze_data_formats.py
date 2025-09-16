import pandas as pd
import re

def analyze_grade_formats():
    """Analyze all the different formats in the grade data"""
    try:
        df = pd.read_excel('KOC Grade Tracker Form(1-52).xlsx', sheet_name='Sheet1')
        
        print("🔍 ANALYZING GRADE DATA FORMATS")
        print("=" * 80)
        
        current_formats = []
        predicted_formats = []
        
        for index, row in df.iterrows():
            name = row['Full Name']
            if pd.isna(name) or name.strip() == '':
                continue
            
            current_text = row['Please list all the subjects you are currently taking and your current grades']
            predicted_text = row['Please list all your predicted grades for each subject']
            
            if not pd.isna(current_text) and str(current_text).strip() not in ['-', 'nan']:
                current_formats.append({
                    'student': name,
                    'text': str(current_text).strip()
                })
            
            if not pd.isna(predicted_text) and str(predicted_text).strip() not in ['-', 'nan']:
                predicted_formats.append({
                    'student': name,
                    'text': str(predicted_text).strip()
                })
        
        print(f"\n📊 CURRENT GRADES FORMATS ({len(current_formats)} samples):")
        print("-" * 60)
        for i, format_example in enumerate(current_formats[:10]):  # Show first 10
            print(f"\n{i+1}. {format_example['student']}:")
            print(f"   '{format_example['text']}'")
        
        print(f"\n🎯 PREDICTED GRADES FORMATS ({len(predicted_formats)} samples):")
        print("-" * 60)
        for i, format_example in enumerate(predicted_formats[:10]):  # Show first 10
            print(f"\n{i+1}. {format_example['student']}:")
            print(f"   '{format_example['text']}'")
        
        # Analyze patterns
        print(f"\n🔍 PATTERN ANALYSIS:")
        print("-" * 60)
        
        patterns_found = {
            'dash_separator': 0,
            'colon_separator': 0,
            'newline_separator': 0,
            'comma_separator': 0,
            'parentheses': 0,
            'mixed_case': 0,
            'all_caps': 0,
            'all_lower': 0,
            'contains_numbers': 0,
            'contains_letters': 0,
            'contains_merit_distinction': 0,
            'contains_gcse_alevel': 0
        }
        
        all_texts = [f['text'] for f in current_formats] + [f['text'] for f in predicted_formats]
        
        for text in all_texts:
            if '-' in text:
                patterns_found['dash_separator'] += 1
            if ':' in text:
                patterns_found['colon_separator'] += 1
            if '\n' in text:
                patterns_found['newline_separator'] += 1
            if ',' in text:
                patterns_found['comma_separator'] += 1
            if '(' in text and ')' in text:
                patterns_found['parentheses'] += 1
            if text.isupper():
                patterns_found['all_caps'] += 1
            elif text.islower():
                patterns_found['all_lower'] += 1
            else:
                patterns_found['mixed_case'] += 1
            if re.search(r'\d', text):
                patterns_found['contains_numbers'] += 1
            if re.search(r'[A-Za-z]', text):
                patterns_found['contains_letters'] += 1
            if re.search(r'(merit|distinction|pass)', text, re.IGNORECASE):
                patterns_found['contains_merit_distinction'] += 1
            if re.search(r'(gcse|a-level|as|btec)', text, re.IGNORECASE):
                patterns_found['contains_gcse_alevel'] += 1
        
        for pattern, count in patterns_found.items():
            percentage = (count / len(all_texts)) * 100 if all_texts else 0
            print(f"   {pattern.replace('_', ' ').title()}: {count} ({percentage:.1f}%)")
        
        return current_formats, predicted_formats
        
    except Exception as e:
        print(f"Error analyzing formats: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    analyze_grade_formats()