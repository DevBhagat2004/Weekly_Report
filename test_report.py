#!/usr/bin/env python3
"""
Test script for the Weekly Report Generator
"""

import os
import sys
import pandas as pd
from datetime import datetime, timedelta
import random

# Add the current directory to Python path to import the main module
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def create_test_data():
    """Create comprehensive test data"""
    print("Creating test data...")
    
    products = ['Laptop Pro', 'Desktop Elite', 'Tablet Max', 'Phone X', 'Watch Smart', 'Headphones Pro']
    regions = ['North America', 'Europe', 'Asia Pacific', 'Latin America', 'Middle East']
    
    data = []
    start_date = datetime.now() - timedelta(days=45)
    
    # Generate 500 records for more comprehensive testing
    for i in range(500):
        # Create dates with higher concentration in recent weeks
        if random.random() < 0.4:  # 40% chance for recent week
            date = datetime.now() - timedelta(days=random.randint(0, 7))
        else:
            date = start_date + timedelta(days=random.randint(0, 45))
        
        product = random.choice(products)
        region = random.choice(regions)
        
        # Create realistic business data
        base_price = {
            'Laptop Pro': 1200,
            'Desktop Elite': 1500,
            'Tablet Max': 800,
            'Phone X': 900,
            'Watch Smart': 400,
            'Headphones Pro': 200
        }[product]
        
        units = random.randint(1, 50)
        price_variation = random.uniform(0.8, 1.2)
        unit_price = base_price * price_variation
        revenue = units * unit_price
        sales_count = random.randint(1, 10)
        
        # Add some data quality issues to test cleaning
        if random.random() < 0.05:  # 5% chance of data issues
            if random.random() < 0.5:
                revenue = f"${revenue:,.2f}"  # Add currency formatting
            else:
                units = f"{units:,}"  # Add comma formatting
        
        data.append({
            'Date': date.strftime('%Y-%m-%d'),
            'Product': product,
            'Sales': sales_count,
            'Units': units,
            'Revenue': round(revenue, 2),
            'Region': region,
            'Customer_Satisfaction': round(random.uniform(3.5, 5.0), 1),
            'Marketing_Spend': round(random.uniform(100, 2000), 2)
        })
    
    # Add some problematic records for testing
    data.append({
        'Date': '2024-12-10',
        'Product': 'Laptop Pro',
        'Sales': '',  # Missing data
        'Units': 10,
        'Revenue': 12000,
        'Region': 'North America',
        'Customer_Satisfaction': 4.2,
        'Marketing_Spend': 500
    })
    
    data.append({
        'Date': 'invalid-date',  # Invalid date
        'Product': 'Phone X',
        'Sales': 5,
        'Units': 15,
        'Revenue': 13500,
        'Region': 'Europe',
        'Customer_Satisfaction': 4.5,
        'Marketing_Spend': 750
    })
    
    # Save test data
    df = pd.DataFrame(data)
    df.to_csv('test_data.csv', index=False)
    print(f"✅ Created test data with {len(data)} records")
    return 'test_data.csv'

def test_individual_functions():
    """Test individual functions of the report generator"""
    print("\n🧪 Testing individual functions...")
    
    try:
        from weekly_report_generator import WeeklyReportGenerator
        
        # Test initialization
        generator = WeeklyReportGenerator()
        print("✅ Generator initialization successful")
        
        # Test data loading
        if generator.load_raw_data('test_data.csv'):
            print("✅ Data loading successful")
            print(f"   Loaded {len(generator.data)} rows")
        else:
            print("❌ Data loading failed")
            return False
        
        # Test data cleaning
        if generator.clean_data():
            print("✅ Data cleaning successful")
            print(f"   Cleaned data: {len(generator.cleaned_data)} rows")
        else:
            print("❌ Data cleaning failed")
            return False
        
        # Test summary generation
        summary = generator.generate_summary_stats()
        if summary:
            print("✅ Summary generation successful")
            print(f"   Generated {len(summary)} metrics")
        else:
            print("❌ Summary generation failed")
            return False
        
        return True
        
    except Exception as e:
        print(f"❌ Error in individual function testing: {e}")
        return False

def test_full_process():
    """Test the complete report generation process"""
    print("\n🔄 Testing full
