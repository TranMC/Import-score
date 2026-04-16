# Test suite for Import Score application
# Install: pip install pytest pandas openpyxl

import pytest
import pandas as pd
import os
import sys

# Add parent directory to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Import functions to test
# Note: Since import_score.py has UI dependencies, we'll test pure functions only
# For full integration tests, we need to mock tkinter components

def test_imports():
    """Test that all required libraries are installed"""
    try:
        import pandas
        import openpyxl
        import re
        from collections import OrderedDict
        assert True
    except ImportError as e:
        pytest.fail(f"Missing required library: {e}")

def test_config_file_exists():
    """Test that config file exists and is valid JSON"""
    config_path = os.path.join(os.path.dirname(__file__), 'app_config.json')
    assert os.path.exists(config_path), "Config file not found"
    
    with open(config_path, 'r', encoding='utf-8') as f:
        import json
        config = json.load(f)
        assert 'excel_reading' in config, "excel_reading section missing in config"
        assert 'header_detection' in config['excel_reading'], "header_detection missing"
        assert 'column_patterns' in config['excel_reading'], "column_patterns missing"

def test_config_structure():
    """Test that config has all required fields"""
    config_path = os.path.join(os.path.dirname(__file__), 'app_config.json')
    with open(config_path, 'r', encoding='utf-8') as f:
        import json
        config = json.load(f)
        
        # Check header_detection
        hd = config['excel_reading']['header_detection']
        assert 'max_search_rows' in hd
        assert 'multi_level_support' in hd
        assert 'merge_separator' in hd
        
        # Check column_patterns
        cp = config['excel_reading']['column_patterns']
        assert 'student_info' in cp
        assert 'score_columns' in cp
        
        # Check student_info patterns
        si = cp['student_info']
        assert 'name' in si
        assert 'patterns' in si['name']
        assert 'regex' in si['name']
        
        # Check score_columns patterns
        sc = cp['score_columns']
        assert 'ddgtx' in sc
        assert 'ddggk' in sc
        assert 'ddgck' in sc

if __name__ == '__main__':
    pytest.main([__file__, '-v'])
