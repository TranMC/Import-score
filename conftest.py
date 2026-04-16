# Pytest fixtures for testing Import Score application

import pytest
import pandas as pd
import os
from datetime import datetime

@pytest.fixture
def sample_student_data():
    """Create sample student data DataFrame"""
    return pd.DataFrame({
        'STT': [1, 2, 3, 4, 5],
        'Họ và tên': ['Nguyễn Văn A', 'Trần Thị B', 'Lê Văn C', 'Phạm Thị D', 'Hoàng Văn E'],
        'Ngày sinh': ['01/01/2005', '02/02/2005', '03/03/2005', '04/04/2005', '05/05/2005'],
        'Giới tính': ['Nam', 'Nữ', 'Nam', 'Nữ', 'Nam'],
        'ĐĐGtx_1': [8, 7, 9, 6, 8],
        'ĐĐGtx_2': [7, 8, 8, 7, 9],
        'ĐĐGtx_3': [9, 7, 8, 8, 7],
        'ĐĐGtx_4': [8, 9, 9, 7, 8],
        'ĐĐGgk': [8.5, 7.5, 9.0, 6.5, 8.0],
        'ĐĐGck': [8.0, 8.0, 9.0, 7.0, 8.5],
        'ĐTBmhkI': [8.2, 7.8, 8.8, 6.8, 8.1]
    })

@pytest.fixture
def sample_excel_file_normal(tmp_path, sample_student_data):
    """Create a sample Excel file with header at row 1 (index 0)"""
    file_path = tmp_path / "test_normal.xlsx"
    sample_student_data.to_excel(file_path, index=False)
    return str(file_path)

@pytest.fixture
def sample_excel_file_header_row_5(tmp_path, sample_student_data):
    """Create a sample Excel file with header at row 5"""
    file_path = tmp_path / "test_header_row_5.xlsx"
    
    # Create DataFrame with empty rows at top
    empty_rows = pd.DataFrame({
        'A': ['UBND QUẬN THANH XUÂN', 'TRƯỜNG THCS KHƯƠNG ĐÌNH', '', 'BẢNG ĐIỂM HỌC KỲ', ''],
        'B': ['', '', '', '', '']
    })
    
    # Combine with actual data
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        # Write title rows
        for idx, row in empty_rows.iterrows():
            pd.DataFrame([row]).to_excel(writer, index=False, header=False, startrow=idx)
        
        # Write actual data starting at row 5
        sample_student_data.to_excel(writer, index=False, startrow=5)
    
    return str(file_path)

@pytest.fixture
def sample_excel_file_large(tmp_path):
    """Create a large Excel file for performance testing"""
    file_path = tmp_path / "test_large.xlsx"
    
    # Create large dataset
    num_students = 1000
    data = {
        'STT': list(range(1, num_students + 1)),
        'Họ và tên': [f'Học sinh {i}' for i in range(1, num_students + 1)],
        'Ngày sinh': ['01/01/2005'] * num_students,
        'Giới tính': ['Nam' if i % 2 == 0 else 'Nữ' for i in range(num_students)],
        'ĐĐGgk': [round(5 + (i % 5) + (i % 10) * 0.1, 1) for i in range(num_students)],
        'ĐĐGck': [round(5 + ((i + 1) % 5) + ((i + 1) % 10) * 0.1, 1) for i in range(num_students)]
    }
    
    df = pd.DataFrame(data)
    df.to_excel(file_path, index=False)
    return str(file_path)

@pytest.fixture
def sample_config():
    """Sample configuration for testing"""
    return {
        "columns": {
            "name": "Tên Học Sinh",
            "exam_code": "Mã đề",
            "score": "Điểm"
        },
        "excel_reading": {
            "header_detection": {
                "enabled": True,
                "max_search_rows": 50,
                "multi_level_support": True,
                "merge_separator": "_",
                "min_columns_match": 3,
                "validate_data_rows": True
            },
            "column_patterns": {
                "student_info": {
                    "name": {
                        "patterns": ["Họ và tên", "Tên học sinh"],
                        "regex": "^(Họ\\s*và\\s*tên|Tên\\s*học\\s*sinh)",
                        "required": True
                    }
                },
                "score_columns": {
                    "ddgtx": {
                        "patterns": ["ĐĐGtx"],
                        "regex": "^(ĐĐGtx)[\\s_-]*(\\d*)$",
                        "has_sub_columns": True
                    }
                }
            }
        }
    }

@pytest.fixture
def empty_dataframe():
    """Empty DataFrame for testing edge cases"""
    return pd.DataFrame()

@pytest.fixture
def dataframe_no_scores(sample_student_data):
    """DataFrame without score columns"""
    df = sample_student_data.copy()
    df = df[['STT', 'Họ và tên', 'Ngày sinh', 'Giới tính']]
    return df
