"""
Cấu hình chung cho Project Report Tool
"""

# Các skill được coi là AI Project
AI_SKILLS = [
    "AI",
    "Machine Learning",
    "ML",
    "Deep Learning",
    "DL",
    "Computer Vision",
    "CV",
    "NLP",
    "Natural Language Processing",
    "Data Science",
    "AI Engineer",
    "ML Engineer",
    "Data Scientist",
]

# Mapping Member Type
MEMBER_TYPE_MAPPING = {
    "Internal": "Internal",
    "X-Jobs": "X-Jobs",
    "Xjobs": "X-Jobs",
    "X-Job": "X-Jobs",
}

# Màu sắc cho định dạng có điều kiện
COLORS = {
    "internal": "C5E0B4",
    "xjobs": "FFC7CE", 
    "ai_project": "FFEB9C",  
    "header_month": "D9E1F2", 
    "header_sub": "F4B084", 
    "summary_row": "FFF2CC", 
    "inactive": "F2F2F2", 
    "fixed_header": "B4C7E7", 
}

# Format số
NUMBER_FORMAT = "#,##0.00"
CURRENCY_FORMAT = "$#,##0.00"
INTEGER_FORMAT = "#,##0"
