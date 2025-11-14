import pandas as pd
import os
import io 
import math
from flask import Flask, jsonify, request, render_template, send_file
from werkzeug.exceptions import NotFound # For handling 404 file errors

# --- 1. Configuration and Setup ---
EXCEL_FILE_PATH = '/tmp/student_records.xlsx'
PASS_CUTOFF = 40 # Pass mark

def init_excel():
    """Ensures the Excel file exists in the Vercel-writable /tmp directory."""
    
    # Check if the file exists OR if it is empty (Vercel resets /tmp frequently)
    # We use a robust check to ensure the file is created if missing or cleared.
    is_missing = not os.path.exists(EXCEL_FILE_PATH)
    is_empty = os.path.exists(EXCEL_FILE_PATH) and os.path.getsize(EXCEL_FILE_PATH) == 0

    if is_missing or is_empty:
        print("Creating new Excel file in /tmp with sample data...")
        
        initial_data = {
            'student_id': ['A001', 'A002', 'A003', 'A004', 'A005'],
            'name': ['Priya Sharma', 'Rajesh Kumar', 'Alia Singh', 'Vikram Taneja', 'Nisha Reddy'],
            'marks': [92, 78, 92, 85, 85]
        }
        df = pd.DataFrame(initial_data)
        
        try:
            # Ensure the directory exists before attempting to write
            os.makedirs(os.path.dirname(EXCEL_FILE_PATH), exist_ok=True)
            df.to_excel(EXCEL_FILE_PATH, index=False)
        except Exception as e:
            # Log the error if file creation fails
            print(f"FATAL: Could not create Excel file in /tmp: {e}")

def read_student_data():
    """Reads all student data from the Excel file into a list of dictionaries."""
    try:
        df = pd.read_excel(EXCEL_FILE_PATH)
        # Ensure marks are treated as integers for correct sorting/statistics
        if 'marks' in df.columns:
             df['marks'] = pd.to_numeric(df['marks'], errors='coerce').fillna(0).astype(int)
        return df.to_dict('records')
    except FileNotFoundError:
        return []
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return []

# Ensure Flask knows where the static files are
app = Flask(__name__, static_folder='static')

# --- 2. Data Structure & Ranking Logic (DS/Algorithm) ---

def calculate_rank(records_sorted_by_marks):
    """Calculates rank correctly, handling ties."""
    if not records_sorted_by_marks: return []
    ranked_list = []
    
    first_student = records_sorted_by_marks[0]
    first_student['rank'] = 1 
    ranked_list.append(first_student)

    for i in range(1, len(records_sorted_by_marks)):
        current_student = records_sorted_by_marks[i]
        prev_student_in_output = ranked_list[-1] 

        if current_student['marks'] == prev_student_in_output['marks']:
            current_rank = prev_student_in_output['rank']
        else:
            current_rank = i + 1
        
        current_student['rank'] = current_rank
        ranked_list.append(current_student)
    return ranked_list

# --- All Sorting Algorithms ---
# (Helper function for Merge Sort)
def merge(left, right, key_name, reverse):
    result = []; i = j = 0
    def compare(a, b):
        if reverse: return a[key_name] >= b[key_name]
        else: return a[key_name] <= b[key_name]
    while i < len(left) and j < len(right):
        if compare(left[i], right[j]): result.append(left[i]); i += 1
        else: result.append(right[j]); j += 1
    result.extend(left[i:]); result.extend(right[j:])
    return result

def merge_sort_students(records, key_name, reverse=False): # O(N log N)
    if len(records) <= 1: return records
    mid = len(records) // 2
    left = merge_sort_students(records[:mid], key_name, reverse)
    right = merge_sort_students(records[mid:], key_name, reverse)
    return merge(left, right, key_name, reverse)

def quick_sort_students(records, key_name, reverse=False): # Avg O(N log N)
    if len(records) <= 1: return records
    pivot = records[len(records) // 2]
    left = []; right = []; equal = []
    for item in records:
        if item[key_name] < pivot[key_name]: left.append(item)
        elif item[key_name] > pivot[key_name]: right.append(item)
        else: equal.append(item)
    sorted_left = quick_sort_students(left, key_name, reverse)
    sorted_right = quick_sort_students(right, key_name, reverse)
    if reverse: return sorted_right + equal + sorted_left
    else: return sorted_left + equal + sorted_right

def bubble_sort_students(records, key_name, reverse=False): # O(N²)
    n = len(records); records_copy = records[:]; swapped = True
    while swapped:
        swapped = False
        for j in range(0, n - 1):
            should_swap = (records_copy[j][key_name] < records_copy[j + 1][key_name]) if reverse else (records_copy[j][key_name] > records_copy[j + 1][key_name])
            if should_swap:
                records_copy[j], records_copy[j + 1] = records_copy[j + 1], records_copy[j]; swapped = True
    return records_copy

def selection_sort_students(records, key_name, reverse=False): # O(N²)
    n = len(records); records_copy = records[:]
    for i in range(n):
        min_max_idx = i
        for j in range(i + 1, n):
            is_min_or_max = (records_copy[j][key_name] > records_copy[min_max_idx][key_name]) if reverse else (records_copy[j][key_name] < records_copy[min_max_idx][key_name])
            if is_min_or_max: min_max_idx = j
        records_copy[i], records_copy[min_max_idx] = records_copy[min_max_idx], records_copy[i]
    return records_copy

def insertion_sort_students(records, key_name, reverse=False): # O(N²)
    n = len(records); records_copy = records[:]
    for i in range(1, n):
        key_item = records_copy[i]; j = i - 1
        while j >= 0:
            should_move = (records_copy[j][key_name] < key_item[key_name]) if reverse else (records_copy[j][key_name] > key_item[key_name])
            if should_move:
                records_copy[j + 1] = records_copy[j]; j -= 1
            else: break
        records_copy[j + 1] = key_item
    return records_copy

# --- 3. API Endpoints ---

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/students')
def get_students():
    records = read_student_data()
    return jsonify(records) 

@app.route('/api/statistics', methods=['GET'])
def get_statistics():
    records = read_student_data()
    marks = [r['marks'] for r in records if 'marks' in r and not pd.isna(r['marks'])]

    if not marks:
        return jsonify({"total": 0, "avg_marks": 0.0, "highest": 0, "lowest": 0, "pass_count": 0, "fail_count": 0, "success": True})
    
    total_students = len(marks)
    highest_marks = max(marks)
    lowest_marks = min(marks)
    avg_marks = sum(marks) / total_students if total_students else 0.0
    pass_count = sum(1 for mark in marks if mark >= PASS_CUTOFF)
    fail_count = total_students - pass_count
            
    return jsonify({
        "total": total_students, "avg_marks": avg_marks, "highest": highest_marks, "lowest": lowest_marks,
        "pass_cutoff": PASS_CUTOFF, "pass_count": pass_count, "fail_count": fail_count, "success": True
    })

@app.route('/api/add_student', methods=['POST'])
def add_student():
    data = request.get_json()
    name = data.get('name'); marks = data.get('marks')
    
    if not name or marks is None:
        return jsonify({"success": False, "message": "Missing name or marks"}), 400

    try:
        df = pd.read_excel(EXCEL_FILE_PATH)
        last_id = df['student_id'].iloc[-1] if not df.empty else 'A000'
        last_num = int(last_id[1:])
        new_id = f"A{last_num + 1:03}"

        new_record = {'student_id': new_id, 'name': name, 'marks': int(marks)}
        df = pd.concat([df, pd.DataFrame([new_record])], ignore_index=True)
        df.to_excel(EXCEL_FILE_PATH, index=False)
        
        return jsonify({"success": True, "message": f"{name} added successfully.", "student_id": new_id})
    except Exception as e:
        return jsonify({"success": False, "message": f"Error saving to Excel: {str(e)}"}), 500

@app.route('/api/upload_excel', methods=['POST'])
def upload_excel():
    if 'file' not in request.files: return jsonify({"success": False, "message": "No file selected."}), 400
    file = request.files['file']
    if file.filename == '': return jsonify({"success": False, "message": "No selected file."}), 400

    try:
        # Save the uploaded file directly to overwrite the existing data file
        file.save(EXCEL_FILE_PATH)
        return jsonify({"success": True, "message": "File uploaded and data updated."})
    except Exception as e:
        return jsonify({"success": False, "message": f"Error processing file: {str(e)}"}), 500

@app.route('/api/generate_sample', methods=['POST'])
def generate_sample_data():
    try:
        initial_data = {
            'student_id': [f'A{i:03}' for i in range(1, 9)],
            'name': ['Alice', 'Bob', 'Charlie', 'David', 'Eve', 'Fiona', 'George', 'Hannah'],
            'marks': [95, 70, 95, 45, 88, 30, 88, 65]
        }
        df = pd.DataFrame(initial_data)
        df.to_excel(EXCEL_FILE_PATH, index=False)
        return jsonify({"success": True, "message": "Sample data generated successfully."})
    except Exception as e:
        return jsonify({"success": False, "message": f"Error generating sample data: {str(e)}"}), 500

@app.route('/api/clear_all', methods=['POST'])
def clear_all_data():
    try:
        df = pd.DataFrame(columns=['student_id', 'name', 'marks'])
        df.to_excel(EXCEL_FILE_PATH, index=False)
        return jsonify({"success": True, "message": "All student data cleared."})
    except Exception as e:
        return jsonify({"success": False, "message": f"Error clearing data: {str(e)}"}), 500

# --- Template Download Route with Fallback ---
@app.route('/static/Rank_Template.xlsx', methods=['GET'])
def download_template_fallback():
    """Tries to send the file; if not found, generates a default template in memory."""
    try:
        # 1. Try to serve the physically stored file first
        return app.send_static_file('Rank_Template.xlsx')
    except NotFound:
        # 2. If file not found, generate a blank template in memory
        df = pd.DataFrame(columns=['student_id', 'name', 'marks'])
        output = io.BytesIO()
        df.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='Rank_Template.xlsx'
        )

@app.route('/api/download_excel', methods=['GET'])
def download_excel():
    """Sorts the data based on user selection, applies rank, and sends the Excel file from memory."""
    sort_by = request.args.get('sort_by', 'marks')
    sort_algo = request.args.get('sort_algorithm', 'built_in')
    reverse_str = request.args.get('reverse', 'true') 
    reverse = reverse_str.lower() == 'true'

    records = read_student_data() 
    
    # Check if records are empty
    if not records:
        df_empty = pd.DataFrame({'status': ['No data available. Please add students or upload a file.']})
        output = io.BytesIO()
        df_empty.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)
        return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='empty_records.xlsx')

    # Algorithm Selection
    sort_functions = {
        'merge_sort': merge_sort_students, 'quick_sort': quick_sort_students,
        'bubble_sort': bubble_sort_students, 'selection_sort': selection_sort_students,
        'insertion_sort': insertion_sort_students,
        'built_in': lambda rec, key, rev: sorted(rec, key=lambda s: s[key], reverse=rev)
    }

    sort_func = sort_functions.get(sort_algo, sort_functions['built_in'])
    final_records = sort_func(records, sort_by, reverse)

    # Apply rank (only if sorting by marks and descending)
    if sort_by == 'marks' and reverse:
        final_records = calculate_rank(final_records)
        
    # Prepare DataFrame for Excel output
    df_sorted = pd.DataFrame(final_records)

    # Ensure rank column is present and properly ordered
    if 'rank' not in df_sorted.columns:
         rank_values = [r.get('rank', 'N/A') for r in final_records]
         df_sorted.insert(0, 'rank', rank_values)

    cols = ['rank', 'student_id', 'name', 'marks']
    df_sorted = df_sorted[[col for col in cols if col in df_sorted.columns]]

    # Create Excel file in memory
    output = io.BytesIO()
    df_sorted.to_excel(output, index=False, engine='openpyxl') 
    output.seek(0) 

    # Send file to browser
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='student_records_ranked.xlsx'
    )

# --- 4. Run Server ---
if __name__ == '__main__':
    init_excel() 
    print("\n\n--- App running! Open http://127.0.0.1:5000/ in your browser ---")
    app.run(debug=True)