# Tests for caching mechanisms

import pytest
import pandas as pd
import os
import sys
import json
from collections import OrderedDict

# Add parent directory to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

class TestConfigCache:
    """Test ConfigCache functionality"""
    
    def test_config_cache_loads_correctly(self, tmp_path):
        """Test that ConfigCache loads config file"""
        from import_score import ConfigCache
        
        # Create test config
        config_data = {"test": "value"}
        config_file = tmp_path / "test_config.json"
        with open(config_file, 'w') as f:
            json.dump(config_data, f)
        
        # Load config
        config = ConfigCache.get_config(str(config_file))
        assert config is not None
        assert config['test'] == 'value'
    
    def test_config_cache_returns_same_instance(self, tmp_path):
        """Test that ConfigCache returns cached instance on second call"""
        from import_score import ConfigCache
        
        config_data = {"test": "value"}
        config_file = tmp_path / "test_config.json"
        with open(config_file, 'w') as f:
            json.dump(config_data, f)
        
        # Load config twice
        config1 = ConfigCache.get_config(str(config_file))
        config2 = ConfigCache.get_config(str(config_file))
        
        # Should be same instance (cached)
        assert config1 is config2
    
    def test_config_cache_invalidation(self):
        """Test that cache can be invalidated"""
        from import_score import ConfigCache
        
        ConfigCache.invalidate()
        assert ConfigCache._cache is None


class TestDataFrameCache:
    """Test DataFrameCache functionality"""
    
    def test_dataframe_cache_stores_and_retrieves(self, tmp_path, sample_student_data):
        """Test storing and retrieving DataFrame from cache"""
        from import_score import DataFrameCache
        
        cache = DataFrameCache(max_size=2)
        file_path = tmp_path / "test.xlsx"
        sample_student_data.to_excel(file_path, index=False)
        
        # Store in cache
        cache.set(str(file_path), sample_student_data)
        
        # Retrieve from cache
        cached_df = cache.get(str(file_path))
        assert cached_df is not None
        assert len(cached_df) == len(sample_student_data)
        assert list(cached_df.columns) == list(sample_student_data.columns)
    
    def test_dataframe_cache_max_size(self, tmp_path, sample_student_data):
        """Test that cache respects max_size"""
        from import_score import DataFrameCache
        
        cache = DataFrameCache(max_size=2)
        
        # Create 3 files
        for i in range(3):
            file_path = tmp_path / f"test_{i}.xlsx"
            sample_student_data.to_excel(file_path, index=False)
            cache.set(str(file_path), sample_student_data)
        
        # First file should be evicted
        first_file = tmp_path / "test_0.xlsx"
        cached_df = cache.get(str(first_file))
        assert cached_df is None  # Should be evicted
    
    def test_dataframe_cache_clear(self, tmp_path, sample_student_data):
        """Test cache clearing"""
        from import_score import DataFrameCache
        
        cache = DataFrameCache()
        file_path = tmp_path / "test.xlsx"
        sample_student_data.to_excel(file_path, index=False)
        
        cache.set(str(file_path), sample_student_data)
        cache.clear()
        
        cached_df = cache.get(str(file_path))
        assert cached_df is None


class TestSearchCache:
    """Test SearchCache functionality"""
    
    def test_search_cache_stores_results(self, sample_student_data):
        """Test storing search results"""
        from import_score import SearchCache
        
        cache = SearchCache()
        cache.clear_if_data_changed(sample_student_data)
        
        # Store search result
        query = "nguyễn"
        result = sample_student_data[sample_student_data['Họ và tên'].str.contains('Nguyễn', case=False, na=False)]
        cache.set(query, result)
        
        # Retrieve
        cached_result = cache.get(query)
        assert cached_result is not None
        assert len(cached_result) == len(result)
    
    def test_search_cache_clears_on_data_change(self, sample_student_data):
        """Test that cache clears when data changes"""
        from import_score import SearchCache
        
        cache = SearchCache()
        cache.clear_if_data_changed(sample_student_data)
        
        # Store result
        cache.set("test", sample_student_data.head(2))
        
        # Change data
        new_df = sample_student_data.copy()
        new_df.loc[0, 'Họ và tên'] = 'Changed'
        
        # Cache should clear
        cache.clear_if_data_changed(new_df)
        result = cache.get("test")
        assert result is None


class TestStatsCache:
    """Test StatsCache functionality"""
    
    def test_stats_cache_calculates_correctly(self, sample_student_data):
        """Test that stats are calculated correctly"""
        from import_score import StatsCache
        
        cache = StatsCache()
        stats = cache.get_stats(sample_student_data)
        
        assert stats['total'] == len(sample_student_data)
        assert 'scored' in stats
        assert 'mean' in stats
        assert 'max' in stats
        assert 'min' in stats
    
    def test_stats_cache_returns_cached_result(self, sample_student_data):
        """Test that stats are cached"""
        from import_score import StatsCache
        
        cache = StatsCache()
        stats1 = cache.get_stats(sample_student_data)
        stats2 = cache.get_stats(sample_student_data)
        
        # Should return same cached result
        assert stats1 == stats2
    
    def test_stats_cache_recalculates_on_change(self, sample_student_data):
        """Test that stats recalculate when data changes"""
        from import_score import StatsCache
        
        cache = StatsCache()
        stats1 = cache.get_stats(sample_student_data)
        
        # Change data
        new_df = sample_student_data.copy()
        new_df = new_df.head(3)  # Reduce rows
        
        stats2 = cache.get_stats(new_df)
        
        assert stats1['total'] != stats2['total']


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
