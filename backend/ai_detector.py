from config import AI_SKILLS


class AIDetector:
    """Class phát hiện dự án AI"""

    def __init__(self, custom_ai_skills=None):
        """
        Khởi tạo AI Detector

        Args:
            custom_ai_skills: Danh sách skill AI tùy chỉnh (optional)
        """
        self.ai_skills = custom_ai_skills if custom_ai_skills else AI_SKILLS
        # Chuyển về lowercase để so sánh không phân biệt hoa thường
        self.ai_skills_lower = [skill.lower() for skill in self.ai_skills]

    def is_ai_project(self, skill):
        """
        Kiểm tra xem skill có phải là AI không

        Args:
            skill: Tên skill cần kiểm tra

        Returns:
            bool: True nếu là AI project, False nếu không
        """
        if not skill or str(skill).strip() == "":
            return False

        skill_lower = str(skill).lower().strip()

        # Kiểm tra exact match
        if skill_lower in self.ai_skills_lower:
            return True

        # Kiểm tra partial match
        for ai_skill in self.ai_skills_lower:
            if ai_skill in skill_lower:
                return True

        return False

    def mark_ai_projects(self, df, skill_column="Skill"):
        """
        Đánh dấu các dự án AI trong DataFrame

        Args:
            df: DataFrame chứa dữ liệu
            skill_column: Tên cột chứa skill

        Returns:
            DataFrame với cột 'AI Project' mới
        """
        # Nếu là AI thì ghi "AI", không thì để trống ""
        df["AI Project"] = df[skill_column].apply(
            lambda x: "AI" if self.is_ai_project(x) else ""
        )
        return df

    def add_ai_skill(self, skill):
        """
        Thêm skill AI mới vào danh sách

        Args:
            skill: Skill cần thêm
        """
        if skill and skill not in self.ai_skills:
            self.ai_skills.append(skill)
            self.ai_skills_lower.append(skill.lower())

    def get_ai_skills_list(self):
        """
        Lấy danh sách các skill AI hiện tại

        Returns:
            list: Danh sách skill AI
        """
        return self.ai_skills.copy()
