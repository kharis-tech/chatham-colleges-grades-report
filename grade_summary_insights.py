import pandas as pd
import re
from collections import defaultdict

def create_executive_summary():
    """Create an executive summary of grade performance"""
    
    print("🎓 UK STUDENT GRADE ANALYSIS - EXECUTIVE SUMMARY")
    print("=" * 80)
    
    # Key findings from the analysis
    students_needing_attention = [
        ("Ricardo Dylan Edwards", "Brompton Academy/Chatham Grammar", "Year 13", 
         ["Applied Science: Pass → Merit", "Sociology: C → A", "BTEC Sport: Merit → Distinction"]),
        
        ("Foday", "St. John Fisher Comprehensive Catholic School", "Year 14",
         ["Applied Science: Pass → Merit", "Health & Social: Merit → Distinction", "Philosophy: D → C"]),
        
        ("Toluwanimi Asaju", "Chatham Grammar", "Year 13",
         ["Business: E → B", "Economics: U → B"]),
        
        ("Elianna Muwanga", "Mayfield Grammar School Gravesend", "Year 10",
         ["History: 7 → 8", "Maths: 4 → 8", "Physics: 7 → 8", "RE: 8 → 9"]),
        
        ("Sophia Nnabue", "Rochester Grammar School for Girls", "Year 10",
         ["English Lit: 6 → 7", "Maths: 5 → 7", "Physics: 6 → 7"])
    ]
    
    high_performers = [
        ("Victoria Dina", "Fort Pitt Grammar School", "Year 12",
         ["Maths: 7 (targeting A*)", "Biology: 8 (targeting A*)", "Chemistry: 8 (targeting A*)"]),
        
        ("Dean Eacott", "Holcombe Grammar School", "Year 13",
         ["Biology: A", "Chemistry: A*", "Psychology: A*"]),
        
        ("Peace Akins", "Brompton Academy", "Year 11",
         ["Meeting all targets across 8+ subjects"])
    ]
    
    print("\n🚨 PRIORITY STUDENTS NEEDING SUPPORT:")
    print("-" * 50)
    for name, school, year, subjects in students_needing_attention:
        print(f"\n👤 {name} ({year})")
        print(f"   🏫 {school}")
        print("   📉 Areas needing improvement:")
        for subject in subjects:
            print(f"      • {subject}")
    
    print("\n\n🌟 HIGH PERFORMING STUDENTS:")
    print("-" * 50)
    for name, school, year, achievements in high_performers:
        print(f"\n⭐ {name} ({year})")
        print(f"   🏫 {school}")
        print("   🎯 Strong performance:")
        for achievement in achievements:
            print(f"      • {achievement}")
    
    print("\n\n📊 KEY INSIGHTS:")
    print("-" * 50)
    print("• Most students are meeting their predicted grades")
    print("• Year 13 students show more variation in performance (exam pressure)")
    print("• BTEC students generally performing well with Merit/Distinction grades")
    print("• Grammar school students have high aspirations (A*/A targets)")
    print("• Some students need additional support in core subjects (Maths, English)")
    
    print("\n\n💡 RECOMMENDATIONS:")
    print("-" * 50)
    print("1. 🎯 Focus intervention on students with multiple subjects below target")
    print("2. 📚 Provide additional Maths support (common area of concern)")
    print("3. 🏆 Celebrate high performers to maintain motivation")
    print("4. 📈 Regular progress reviews for Year 13 students approaching exams")
    print("5. 🤝 Peer mentoring between high performers and struggling students")
    
    print("\n" + "=" * 80)

if __name__ == "__main__":
    create_executive_summary()