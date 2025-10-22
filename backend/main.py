"""
Module chính để chạy Project Report Tool
"""

import os
import sys
from datetime import datetime
from data_processor import DataProcessor
from ai_detector import AIDetector
from calculator import RevenueCalculator
from report_generator import ReportGenerator
import pandas as pd


class ProjectReportTool:
    """Class chính điều phối toàn bộ quy trình"""

    def __init__(self):
        self.data_processor = DataProcessor()
        self.ai_detector = AIDetector()
        self.calculator = RevenueCalculator()
        self.report_generator = ReportGenerator()

    def load_project_code_file(self, file_path):
        """
        Đọc file project_code.xlsx để lấy mapping Revenue

        Args:
            file_path: Đường dẫn file project_code.xlsx

        Returns:
            tuple: (DataFrame gốc, dict mapping {Project Code: Ratecard})
        """
        try:
            # Đọc file Excel
            if file_path.endswith(".xls"):
                df = pd.read_excel(file_path, engine="xlrd")
            else:
                df = pd.read_excel(file_path, engine="openpyxl")

            # Chuẩn hóa tên cột
            df.columns = df.columns.str.strip()

            # Kiểm tra các cột bắt buộc
            if "Project Code" not in df.columns:
                raise Exception("File project_code.xlsx thiếu cột 'Project Code'")
            if "Ratecard" not in df.columns:
                raise Exception("File project_code.xlsx thiếu cột 'Ratecard'")

            # Tạo mapping dictionary
            revenue_mapping = {}
            for _, row in df.iterrows():
                project_code = str(row["Project Code"]).strip()
                ratecard = row["Ratecard"]

                # Xử lý giá trị ratecard
                try:
                    ratecard_value = float(ratecard) if pd.notna(ratecard) else 0
                except (ValueError, TypeError):
                    ratecard_value = 0

                revenue_mapping[project_code] = ratecard_value

            print(f"✓ Đã load {len(revenue_mapping)} project codes từ file")

            return df, revenue_mapping

        except Exception as e:
            raise Exception(f"Lỗi đọc file project_code.xlsx: {str(e)}")

    def select_date_range(self, available_months):
        """
        Cho phép user chọn khoảng thời gian để XUẤT (filter data)

        Args:
            available_months: Danh sách (year, month) có sẵn

        Returns:
            tuple: ((start_year, start_month), (end_year, end_month)) hoặc (None, None) nếu xuất tất cả
        """
        print("\n" + "=" * 70)
        print("LỰA CHỌN KHOẢNG THỜI GIAN")
        print("=" * 70)
        print("\nCác tháng có sẵn trong dữ liệu:")

        for idx, (year, month) in enumerate(available_months, 1):
            month_name = self.report_generator.get_month_name(month)
            print(f"  {idx}. {month_name} {year}")

        print("\n💡 Nhập 'all' hoặc để trống để xuất TẤT CẢ các tháng")
        print("   Hoặc nhập số thứ tự để chọn khoảng thời gian cụ thể")

        choice = input("\nLựa chọn: ").strip().lower()

        if choice == "" or choice == "all":
            print("✓ Sẽ xuất tất cả các tháng")
            return None, None

        # Chọn tháng bắt đầu
        try:
            start_idx = int(choice)
            if not (1 <= start_idx <= len(available_months)):
                print("  ✖ Số không hợp lệ! Sẽ xuất tất cả.")
                return None, None
        except ValueError:
            print("  ✖ Giá trị không hợp lệ! Sẽ xuất tất cả.")
            return None, None

        # Chọn tháng kết thúc
        while True:
            try:
                end_idx = int(
                    input(
                        f"Chọn tháng KẾT THÚC ({start_idx}-{len(available_months)}): "
                    )
                )
                if start_idx <= end_idx <= len(available_months):
                    break
                print(
                    f"  ✖ Vui lòng nhập số từ {start_idx} đến {len(available_months)}"
                )
            except ValueError:
                print("  ✖ Vui lòng nhập số!")

        start_month = available_months[start_idx - 1]
        end_month = available_months[end_idx - 1]

        month_name_start = self.report_generator.get_month_name(start_month[1])
        month_name_end = self.report_generator.get_month_name(end_month[1])

        print(
            f"\n✓ Đã chọn: {month_name_start} {start_month[0]} đến {month_name_end} {end_month[0]}"
        )

        return start_month, end_month

    def select_revenue_month(self, available_months):
        """
        Cho phép user chọn tháng để tạo Revenue_By_Account sheet

        Args:
            available_months: Danh sách (year, month) có sẵn

        Returns:
            tuple: (year, month) hoặc None
        """
        print("\n" + "=" * 70)
        print("CHỌN THÁNG CHO BÁO CÁO REVENUE BY ACCOUNT")
        print("=" * 70)
        print("\nCác tháng có sẵn:")

        for idx, (year, month) in enumerate(available_months, 1):
            month_name = self.report_generator.get_month_name(month)
            print(f"  {idx}. {month_name} {year}")

        print("\nNhập 'skip' hoặc để trống để bỏ qua sheet này")

        choice = input("\nLựa chọn tháng: ").strip().lower()

        if choice == "" or choice == "skip":
            return None

        try:
            month_idx = int(choice)
            if 1 <= month_idx <= len(available_months):
                selected = available_months[month_idx - 1]
                month_name = self.report_generator.get_month_name(selected[1])
                print(f"\n✓ Đã chọn: {month_name} {selected[0]}")
                return selected
            else:
                print("  ✖ Số không hợp lệ! Bỏ qua Revenue_By_Account sheet.")
                return None
        except ValueError:
            print("  ✖ Giá trị không hợp lệ! Bỏ qua Revenue_By_Account sheet.")
            return None

    def run(self, input_file, project_code_file, output_file=None):
        """
        Chạy toàn bộ quy trình tạo báo cáo

        Args:
            input_file: Đường dẫn file Excel đầu vào (data)
            project_code_file: Đường dẫn file project_code.xlsx
            output_file: Đường dẫn file Excel đầu ra (optional)
        """
        try:
            print("=" * 70)
            print("PROJECT REPORT TOOL")
            print("=" * 70)

            # 1. Đọc file project_code.xlsx
            print("\n[1/8] Đang đọc file project_code.xlsx...")
            df_project_code, revenue_mapping = self.load_project_code_file(
                project_code_file
            )

            # 2. Đọc dữ liệu đầu vào
            print("\n[2/8] Đang đọc file đầu vào...")
            df_input = self.data_processor.load_data(input_file)
            print(f"✓ Đã đọc {len(df_input)} dòng dữ liệu")

            # 3. Lấy danh sách Project Code duy nhất
            print("\n[3/8] Phát hiện Project Codes...")
            project_codes = self.data_processor.get_unique_project_codes(df_input)
            print(f"✓ Tìm thấy {len(project_codes)} project codes:")
            for i, pc in enumerate(project_codes, 1):
                rev = revenue_mapping.get(pc, 0)
                print(f"  {i}. {pc} → Revenue: {rev}")

            # Kiểm tra project codes thiếu
            missing_codes = [pc for pc in project_codes if pc not in revenue_mapping]
            if missing_codes:
                print(
                    "\n⚠ Cảnh báo: Các project code sau không có trong file project_code.xlsx:"
                )
                for pc in missing_codes:
                    print(f"  - {pc} (sẽ dùng revenue = 0)")

            # 4. Thêm Revenue vào DataFrame
            print("\n[4/8] Đang áp dụng Revenue vào dữ liệu...")
            df_input = self.data_processor.add_revenue_to_data(
                df_input, revenue_mapping
            )
            print("✓ Đã thêm Revenue cho tất cả dòng dữ liệu")

            # 5. Phân bổ dữ liệu theo tháng (chỉ để tính Summary)
            print("\n[5/8] Đang phân bổ dữ liệu theo tháng...")
            df_monthly = self.data_processor.allocate_by_month(df_input)
            available_months = self.data_processor.get_available_months(df_monthly)
            print(f"✓ Đã phân bổ dữ liệu cho {len(available_months)} tháng")

            # 6. Đánh dấu AI projects
            print("\n[6/8] Đang nhận diện AI projects...")
            df_input = self.ai_detector.mark_ai_projects(df_input)
            df_monthly = self.ai_detector.mark_ai_projects(df_monthly)
            ai_count_input = len(df_input[df_input["AI Project"] == "AI"])
            ai_count_monthly = len(df_monthly[df_monthly["AI Project"] == "AI"])
            print(f"✓ Input: {ai_count_input} dòng AI projects")
            print(f"✓ Monthly: {ai_count_monthly} dòng AI projects")

            # 7. Thêm MAIL column vào input
            print("\n[7/8] Chuẩn bị dữ liệu...")
            df_input["MAIL"] = df_input["Username"].apply(lambda x: f"{x}@fpt.com")

            # 8. Tính toán cho Summary sheet (từ df_monthly)
            print("\n[8/8] Tính toán metrics cho Summary sheet...")
            df_monthly = self.calculator.add_calculations(df_monthly)

            # Hiển thị thống kê
            stats = self.calculator.get_summary_statistics(df_monthly)
            print("\nThống kê:")
            print(f"  - Tổng records (input): {len(df_input)}")
            print(f"  - Tổng records (monthly): {len(df_monthly)}")
            print(f"  - Unique users: {stats['unique_users']}")
            print(f"  - Unique projects: {stats['unique_projects']}")
            print(f"  - Internal: {stats['internal_count']} records")
            print(f"  - X-Jobs: {stats['xjobs_count']} records")
            print(f"  - AI Projects: {stats['ai_projects_count']} records")
            print(f"  - Total Revenue: ${stats['total_revenue']:,.2f}")
            print(f"  - Total AI Revenue: ${stats['total_ai_revenue']:,.2f}")

            # 9. Tạo file output
            if output_file is None:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_file = f"backend/output/report_{timestamp}.xlsx"

            os.makedirs(os.path.dirname(output_file), exist_ok=True)

            # 10. Tạo báo cáo 2 sheets
            print("\nTạo báo cáo Excel...")
            self.report_generator.generate_report_two_sheets(
                df_input=df_input,  # 95 records gốc
                df_monthly=df_monthly,  # 472 records allocate
                month_list=available_months,
                output_path=output_file,
                df_project_code=df_project_code,
            )

            print(f"\n✓ Báo cáo đã được lưu tại: {output_file}")
            print(f"  - Sheet 1 (Project Report): {len(df_input)} rows")
            print("  - Sheet 2 (Summary): Monthly metrics")

            print("\n" + "=" * 70)
            print("HOÀN THÀNH!")
            print("=" * 70)

            return output_file
            print("HOÀN THÀNH!")
            print("=" * 70)
            print(
                "\n💡 Lưu ý: Bạn có thể sửa giá trị Ratecard trong sheet 'Project_Code'"
            )
            print("   và các sheet khác sẽ tự động cập nhật theo!")

            return output_file

        except Exception as e:
            print(f"\n✖ LỖI: {str(e)}")
            import traceback

            traceback.print_exc()
            sys.exit(1)

    def validate_input_file(self, file_path):
        """
        Kiểm tra tính hợp lệ của file đầu vào

        Args:
            file_path: Đường dẫn file cần kiểm tra

        Returns:
            bool: True nếu hợp lệ
        """
        if not os.path.exists(file_path):
            print(f"✖ File không tồn tại: {file_path}")
            return False

        if not (
            file_path.lower().endswith(".xls") or file_path.lower().endswith(".xlsx")
        ):
            print("✖ File phải có định dạng .xls hoặc .xlsx")
            return False

        # Kiểm tra các cột bắt buộc
        try:
            if file_path.endswith(".xls"):
                df = pd.read_excel(file_path, nrows=1, engine="xlrd")
            else:
                df = pd.read_excel(file_path, nrows=1, engine="openpyxl")

            required_columns = [
                "Username",
                "Project Code",
                "From Date",
                "To Date",
                "Member Type",
                "Calendar Effort",
                "Skill",
            ]

            missing_columns = []
            for col in required_columns:
                if col not in df.columns:
                    missing_columns.append(col)

            if missing_columns:
                print(f"✖ Thiếu các cột bắt buộc: {', '.join(missing_columns)}")
                return False

            return True

        except Exception as e:
            print(f"✖ Lỗi khi đọc file: {str(e)}")
            return False


def main():
    """Hàm main để chạy từ command line"""

    if len(sys.argv) < 3:
        print("Cách sử dụng:")
        print(
            "  python main.py <input_file.xls> <project_code.xlsx> [output_file.xlsx]"
        )
        print("\nVí dụ:")
        print(
            "  python main.py data/input/sample_input.xls data/input/project_code.xlsx"
        )
        print(
            "  python main.py data/input/sample_input.xls data/input/project_code.xlsx backend/output/my_report.xlsx"
        )
        sys.exit(1)

    input_file = sys.argv[1]
    project_code_file = sys.argv[2]
    output_file = sys.argv[3] if len(sys.argv) > 3 else None

    # Khởi tạo tool
    tool = ProjectReportTool()

    # Validate input files
    if not tool.validate_input_file(input_file):
        sys.exit(1)

    if not os.path.exists(project_code_file):
        print(f"File project_code.xlsx không tồn tại: {project_code_file}")
        sys.exit(1)

    # Chạy tool
    tool.run(input_file, project_code_file, output_file)


if __name__ == "__main__":
    main()
