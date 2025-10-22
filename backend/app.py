"""
Flask API Server cho Project Report Tool
"""

from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os
from werkzeug.utils import secure_filename
import tempfile
import shutil

from data_processor import DataProcessor
from ai_detector import AIDetector
from calculator import RevenueCalculator
from report_generator import ReportGenerator
from main import ProjectReportTool
from main_multi_files import MultiFileProjectReportTool

app = Flask(__name__)
CORS(app)  # Enable CORS cho React

# Cấu hình
UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
ALLOWED_EXTENSIONS = {"xls", "xlsx"}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["OUTPUT_FOLDER"] = OUTPUT_FOLDER
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB max


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route("/api/health", methods=["GET"])
def health_check():
    """Health check endpoint"""
    return jsonify({"status": "ok", "message": "Server is running"})


@app.route("/api/process/single", methods=["POST"])
def process_single_file():
    """
    Xử lý single file mode

    Form data:
    - project_code: file project_code.xlsx
    - input_file: file input data
    """
    try:
        # Validate files
        if "project_code" not in request.files:
            return jsonify({"error": "Missing project_code file"}), 400
        if "input_file" not in request.files:
            return jsonify({"error": "Missing input_file"}), 400

        project_code_file = request.files["project_code"]
        input_file = request.files["input_file"]

        if not allowed_file(project_code_file.filename):
            return jsonify({"error": "Invalid project_code file format"}), 400
        if not allowed_file(input_file.filename):
            return jsonify({"error": "Invalid input file format"}), 400

        # Save files
        pc_filename = secure_filename(project_code_file.filename)
        input_filename = secure_filename(input_file.filename)

        pc_path = os.path.join(app.config["UPLOAD_FOLDER"], pc_filename)
        input_path = os.path.join(app.config["UPLOAD_FOLDER"], input_filename)

        project_code_file.save(pc_path)
        input_file.save(input_path)

        # Process
        tool = ProjectReportTool()
        output_filename = f"report_{os.urandom(8).hex()}.xlsx"
        output_path = os.path.join(app.config["OUTPUT_FOLDER"], output_filename)

        tool.run(input_path, pc_path, output_path)

        # Cleanup uploaded files
        os.remove(pc_path)
        os.remove(input_path)

        return jsonify(
            {
                "success": True,
                "output_file": output_filename,
                "message": "Processing completed successfully",
            }
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/process/multi", methods=["POST"])
def process_multi_files():
    """
    Xử lý multi files mode

    Form data:
    - project_code: file project_code.xlsx
    - files[]: array of input files
    - metadata: JSON string với [{"filename": "...", "year": 2024, "month": 1}, ...]
    """
    try:
        # Validate
        if "project_code" not in request.files:
            return jsonify({"error": "Missing project_code file"}), 400

        project_code_file = request.files["project_code"]

        # Get files
        files = request.files.getlist("files[]")
        if not files or len(files) == 0:
            return jsonify({"error": "No input files provided"}), 400

        # Get metadata
        import json

        metadata = json.loads(request.form.get("metadata", "[]"))

        if len(files) != len(metadata):
            return jsonify({"error": "Files and metadata count mismatch"}), 400

        # Save project code
        pc_filename = secure_filename(project_code_file.filename)
        pc_path = os.path.join(app.config["UPLOAD_FOLDER"], pc_filename)
        project_code_file.save(pc_path)

        # Save all input files
        file_list = []
        for file, meta in zip(files, metadata):
            if not allowed_file(file.filename):
                continue

            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
            file.save(file_path)

            file_list.append(
                (file_path, int(meta.get("year", 2024)), int(meta.get("month", 1)))
            )

        # Process
        tool = MultiFileProjectReportTool()
        output_filename = f"merged_report_{os.urandom(8).hex()}.xlsx"
        output_path = os.path.join(app.config["OUTPUT_FOLDER"], output_filename)

        tool.run_multi_files(file_list, pc_path, output_path)

        # Cleanup
        os.remove(pc_path)
        for file_path, _, _ in file_list:
            if os.path.exists(file_path):
                os.remove(file_path)

        return jsonify(
            {
                "success": True,
                "output_file": output_filename,
                "message": "Processing completed successfully",
                "files_processed": len(file_list),
            }
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/download/<filename>", methods=["GET"])
def download_file(filename):
    """Download generated report"""
    try:
        file_path = os.path.join(app.config["OUTPUT_FOLDER"], filename)

        if not os.path.exists(file_path):
            return jsonify({"error": "File not found"}), 404

        return send_file(
            file_path,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
