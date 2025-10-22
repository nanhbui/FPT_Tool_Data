import pandas as pd

# from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from config import MEMBER_TYPE_MAPPING


class DataProcessor:
    """Class xử lý và phân bổ dữ liệu"""

    def __init__(self):
        self.month_columns = []

    def load_data(self, file_path):
        """
        Đọc file Excel đầu vào (.xls hoặc .xlsx)

        Args:
            file_path: Đường dẫn file Excel

        Returns:
            DataFrame chứa dữ liệu
        """
        try:
            # Xác định engine dựa trên extension
            if file_path.endswith(".xls"):
                df = pd.read_excel(file_path, engine="xlrd")
            else:
                df = pd.read_excel(file_path, engine="openpyxl")

            # Chuẩn hóa tên cột
            df.columns = df.columns.str.strip()
            return df
        except Exception as e:
            raise Exception(f"Lỗi đọc file: {str(e)}")

    def get_unique_project_codes(self, df):
        """
        Lấy danh sách Project Code duy nhất

        Args:
            df: DataFrame chứa dữ liệu

        Returns:
            list: Danh sách Project Code unique
        """
        project_codes = df["Project Code"].unique()
        return sorted(project_codes)

    def add_revenue_to_data(self, df, revenue_mapping):
        """
        Thêm Revenue vào DataFrame dựa trên mapping

        Args:
            df: DataFrame gốc
            revenue_mapping: Dict {Project Code: Revenue}

        Returns:
            DataFrame với cột Revenue
        """
        df["Revenue"] = df["Project Code"].map(revenue_mapping)
        return df

    def normalize_member_type(self, member_type):
        """
        Chuẩn hóa Member Type

        Args:
            member_type: Giá trị Member Type gốc

        Returns:
            str: Giá trị đã chuẩn hóa
        """
        if not member_type:
            return "Internal"

        member_type_str = str(member_type).strip()
        return MEMBER_TYPE_MAPPING.get(member_type_str, "Internal")

    def get_months_between(self, from_date, to_date):
        """
        Lấy danh sách các tháng trong khoảng thời gian

        Args:
            from_date: Ngày bắt đầu
            to_date: Ngày kết thúc

        Returns:
            list: Danh sách (year, month)
        """
        months = []
        current = from_date.replace(day=1)
        end = to_date.replace(day=1)

        while current <= end:
            months.append((current.year, current.month))
            current += relativedelta(months=1)

        return months

    def allocate_by_month(self, df):
        """
        Phân bổ dữ liệu theo từng tháng

        Args:
            df: DataFrame đầu vào

        Returns:
            DataFrame đã được phân bổ theo tháng
        """
        # Chuyển đổi cột ngày tháng
        df["From Date"] = pd.to_datetime(df["From Date"])
        df["To Date"] = pd.to_datetime(df["To Date"])

        # Chuẩn hóa Member Type
        df["Member Type"] = df["Member Type"].apply(self.normalize_member_type)

        # Tạo MAIL từ Username
        df["MAIL"] = df["Username"].apply(lambda x: f"{x}@fpt.com")

        # Phân bổ theo tháng
        monthly_data = []

        for _, row in df.iterrows():
            months = self.get_months_between(row["From Date"], row["To Date"])

            for year, month in months:
                # LẤY TRỰC TIẾP Calendar Effort từ input, KHÔNG tính toán lại
                monthly_data.append(
                    {
                        "Username": row["Username"],
                        "MAIL": row["MAIL"],
                        "Project Code": row["Project Code"],
                        "Member Type": row["Member Type"],
                        "Revenue": row["Revenue"],
                        "Skill": row["Skill"],
                        "Year": year,
                        "Month": month,
                        "Calendar Effort": row["Calendar Effort"],  # Lấy trực tiếp
                    }
                )

        return pd.DataFrame(monthly_data)

    def filter_by_date_range(self, df, start_year, start_month, end_year, end_month):
        """
        Lọc dữ liệu theo khoảng thời gian

        Args:
            df: DataFrame đã phân bổ theo tháng
            start_year: Năm bắt đầu
            start_month: Tháng bắt đầu
            end_year: Năm kết thúc
            end_month: Tháng kết thúc

        Returns:
            DataFrame đã lọc
        """
        mask = (
            ((df["Year"] == start_year) & (df["Month"] >= start_month))
            | ((df["Year"] > start_year) & (df["Year"] < end_year))
            | ((df["Year"] == end_year) & (df["Month"] <= end_month))
        )
        return df[mask].copy()

    def get_unique_months(self, df):
        """
        Lấy danh sách các tháng duy nhất trong dữ liệu

        Args:
            df: DataFrame đã phân bổ theo tháng

        Returns:
            list: Danh sách (year, month) đã sắp xếp
        """
        months = df[["Year", "Month"]].drop_duplicates()
        months = months.sort_values(["Year", "Month"])
        self.month_columns = [
            (row["Year"], row["Month"]) for _, row in months.iterrows()
        ]
        return self.month_columns

    def get_available_months(self, df):
        """
        Lấy danh sách tháng có sẵn trong dữ liệu (để user chọn)

        Args:
            df: DataFrame đã phân bổ theo tháng

        Returns:
            list: Danh sách (year, month) đã sắp xếp
        """
        return self.get_unique_months(df)
