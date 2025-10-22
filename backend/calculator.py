class RevenueCalculator:
    """Class tính toán revenue"""

    def calculate_rev_eff(self, revenue, effort):
        """
        Tính Revenue × Calendar Effort

        Args:
            revenue: Giá trị revenue
            effort: Calendar effort

        Returns:
            float: Kết quả tính toán
        """
        try:
            return round(float(revenue) * float(effort), 2)
        except (ValueError, TypeError):
            return 0.0

    def calculate_ai_rev(self, revenue, effort, is_ai):
        """
        Tính AI Revenue (chỉ cho AI projects)

        Args:
            revenue: Giá trị revenue
            effort: Calendar effort
            is_ai: True nếu là AI project

        Returns:
            float: Kết quả tính toán (0 nếu không phải AI)
        """
        if is_ai:
            return self.calculate_rev_eff(revenue, effort)
        return 0.0

    def add_calculations(self, df):
        """
        Thêm các cột tính toán vào DataFrame

        Args:
            df: DataFrame với dữ liệu monthly

        Returns:
            DataFrame với các cột tính toán mới
        """
        # Tính REVxEFF (cho tất cả projects)
        df["REVxEFF"] = df.apply(
            lambda row: self.calculate_rev_eff(row["Revenue"], row["Calendar Effort"]),
            axis=1,
        )

        # Tính AI-REV (chỉ cho AI projects)
        # Kiểm tra cột AI Project == "AI" (không phải "Non-AI" nữa)
        df["AI-REV"] = df.apply(
            lambda row: self.calculate_ai_rev(
                row["Revenue"],
                row["Calendar Effort"],
                row.get("AI Project", "") == "AI",
            ),
            axis=1,
        )

        return df

    def aggregate_by_user_project_month(self, df):
        """
        Tổng hợp dữ liệu theo User, Project, và Month

        Args:
            df: DataFrame với dữ liệu chi tiết

        Returns:
            DataFrame đã tổng hợp
        """
        group_cols = [
            "Username",
            "MAIL",
            "Project Code",
            "Member Type",
            "Revenue",
            "AI Project",
            "Year",
            "Month",
        ]

        agg_dict = {"Calendar Effort": "sum", "REVxEFF": "sum", "AI-REV": "sum"}

        result = df.groupby(group_cols, as_index=False).agg(agg_dict)

        # Làm tròn các giá trị
        result["Calendar Effort"] = result["Calendar Effort"].round(2)
        result["REVxEFF"] = result["REVxEFF"].round(2)
        result["AI-REV"] = result["AI-REV"].round(2)

        return result

    def get_summary_statistics(self, df):
        """
        Tính toán thống kê tổng hợp

        Args:
            df: DataFrame với dữ liệu

        Returns:
            dict: Thống kê tổng hợp
        """
        stats = {
            "total_effort": df["Calendar Effort"].sum(),
            "total_revenue": df["REVxEFF"].sum(),
            "total_ai_revenue": df["AI-REV"].sum(),
            "internal_count": len(df[df["Member Type"] == "Internal"]),
            "xjobs_count": len(df[df["Member Type"] == "X-Jobs"]),
            "ai_projects_count": len(df[df["AI Project"] == "AI"]),
            "unique_users": df["Username"].nunique(),
            "unique_projects": df["Project Code"].nunique(),
        }

        return stats
