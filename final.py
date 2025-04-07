import pandas as pd
import os
import numpy as np
import json
import io
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg') # Sử dụng Agg backend để tránh lỗi trên server không có GUI
import seaborn as sns
from matplotlib.backends.backend_pdf import PdfPages
from fpdf import FPDF
from datetime import datetime
import base64
import traceback
from flask import request, jsonify


def load_data():
    base_path = 'C:/Users/Dell/OneDrive/Tài liệu/GÓI 1/127 hoan chinh/data'
    files = {
        'avg_by_code': 'Average_by_Code.csv',
        'avg_by_sector': 'Average_by_Sector.csv',
        'balance_sheet': 'BCDKT.csv',
        'fin_statements': 'BCTC.csv',
        'income_statement': 'KQKD.csv',
        'cash_flow': 'LCTT.csv',
        'disclosures': 'TM.csv'
    }
    data = {}
    for key, filename in files.items():
        path = os.path.join(base_path, filename)
        try:
            pass  # Replace with the intended logic or remove the try block if unnecessary
        except Exception as e:
            print(f"An error occurred: {e}")
        except Exception as e:
            print(f"An error occurred: {e}")
        except Exception as e:
            print(f"An error occurred: {e}")
            data[key] = pd.read_csv(path, encoding='utf-8')
        except Exception as e:
            print(f"Error loading {filename}: {e}")
            data[key] = pd.DataFrame() # Use an empty DataFrame as a fallback
    
    # Load Excel file
    excel_path = os.path.join(base_path, 'thongtin.xlsx')
    try:
        data['company_info'] = pd.read_excel(excel_path)
        print("Đã load thongtin.xlsx")
    except Exception as e:
        print(f"Lỗi khi load thongtin.xlsx: {e}")
        data['company_info'] = pd.DataFrame()
    
    return data

data = load_data()
fin_statements = data['fin_statements']

# Hàm lấy dữ liệu cho báo cáo
def get_company_report_data(company_code):
    result = {} # Khởi tạo result
    try:
        # Logic xử lý dữ liệu
        result['company_code'] = company_code
        result['report_date'] = datetime.now().strftime('%d/%m/%Y')
        
        company_info = fin_statements[fin_statements['Mã'] == company_code]
        if company_info.empty:
            print(f"Không tìm thấy mã {company_code} trong BCTC.csv")
            return None
        
        # Lấy thông tin công ty
        result['company_name'] = company_info['Tên công ty'].iloc[3]
        result['exchange'] = company_info['Sàn'].iloc[4] if 'Sàn' in company_info.columns else "N/A"
        result['industry_level1'] = company_info['Ngành ICB - cấp 1'].iloc[5] if 'Ngành ICB - cấp 1' in company_info.columns else "N/A"
        result['industry_level2'] = company_info['Ngành ICB - cấp 2'].iloc[6] if 'Ngành ICB - cấp 2' in company_info.columns else "N/A"
        result['industry_level3'] = company_info['Ngành ICB - cấp 3'].iloc[7] if 'Ngành ICB - cấp 3' in company_info.columns else "N/A"
        
        # Lọc theo mã
        bs = data['balance_sheet'][data['balance_sheet']['Mã'] == company_code].sort_values(by=['Năm', 'Quý']) 
        income_statement = data['income_statement']
        is_ = income_statement[income_statement['Mã'] == company_code].sort_values(by=['Năm', 'Quý'])
        cf = data['cash_flow'][data['cash_flow']['Mã'] == company_code].sort_values(by=['Năm', 'Quý'])
        
        # Trung bình ngành
        sector_avg = {}
        if result['industry_level3'] != "N/A": 
            avg_by_sector = data.get('avg_by_sector', pd.DataFrame()) # Ensure avg_by_sector is loaded
            sector_data = avg_by_sector[avg_by_sector['Sector'] == result['industry_level3']]
            if not sector_data.empty:
                sector_avg = sector_data.iloc[0].to_dict()
                result['sector_avg'] = sector_avg
        
        # Chỉ số riêng công ty
        avg_by_code = data.get('avg_by_code', pd.DataFrame()) # Ensure avg_by_code is loaded
        company_metrics = avg_by_code[avg_by_code['Mã'] == company_code]
        if not company_metrics.empty:
            result['company_metrics'] = company_metrics.iloc[0].to_dict()
        
        # Lấy các năm có dữ liệu
        years = []
        for df in [bs, is_, cf]:
            if 'Năm' in df.columns:
                years.extend(df['Năm'].dropna().astype(int).unique()) 
        years = sorted(list(set(years)))
        if not years:
            print(f"Không có dữ liệu tài chính theo năm cho mã {company_code}")
            result['years'] = []
            result['financial_data'] = {}
            result['financial_ratios'] = {}
            # Ensure this statement is inside a function or remove it if unnecessary
            pass # Replace with appropriate logic if needed
        
        # Giới hạn 5 năm gần nhất
        years = years[-5:]
        financial_data = {}
        
        for year in years:
            year_data = {
                'balance_sheet': {},
                'income_statement': {},
                'cash_flow': {},
            }
            
            # Lấy dữ liệu cuối quý của năm
            def get_last_row(df):
                year_df = df[df['Năm'] == year].sort_values(by='Quý', ascending=False)
                return year_df.iloc[0] if not year_df.empty else {}
            bs_row = get_last_row(bs)
            is_row = get_last_row(is_)
            cf_row = get_last_row(cf)
            
            # Mapping trường
            bs_fields = {
                'total_assets': 'TỔNG CỘNG TÀI SẢN',
                'current_assets': 'TÀI SẢN NGẮN HẠN',
                'fixed_assets': 'TÀI SẢN DÀI HẠN',
                'liabilities': 'NỢ PHẢI TRẢ',
                'equity': 'VỐN CHỦ SỞ HỮU',
                'inventory': 'Hàng tồn kho, ròng',
                'short_term_debt': 'Nợ ngắn hạn'
            }
            
            is_fields = {
                'revenue': 'Doanh thu thuần',
                'gross_profit': 'Lợi nhuận gộp về bán hàng và cung cấp dịch vụ',
                'operating_profit': 'Lợi nhuận thuần từ hoạt động kinh doanh',
                'profit_before_tax': 'Tổng lợi nhuận kế toán trước thuế',
                'net_profit': 'Lợi nhuận sau thuế thu nhập doanh nghiệp',
                'interest_expense': 'Trong đó: Chi phí lãi vay',
            }

            for k, v in is_fields.items():
                year_data['income_statement'][k] = is_row.get(v, 0)
                
            cf_fields = {
                'operating_cash_flow': 'Lưu chuyển tiền từ hoạt động kinh doanh',
                'investing_cash_flow': 'Lưu chuyển tiền từ hoạt động đầu tư',
                'financing_cash_flow': 'Lưu chuyển tiền từ hoạt động tài chính',
            }

            for k, v in cf_fields.items():
                year_data['cash_flow'][k] = cf_row.get(v, 0)
                
            financial_data[year] = year_data
        
        result['years'] = years
        result['financial_data'] = financial_data
        
        # Calculate EBIT (Earnings Before Interest and Taxes)
        interest_expense = year_data['income_statement']['interest_expense']
        profit_before_tax = year_data['income_statement']['profit_before_tax']
        if interest_expense > 0 or profit_before_tax > 0: 
            year_data['income_statement']['ebit'] = profit_before_tax + interest_expense
            
        # Calculate EBITDA (EBIT + Depreciation & Amortization)
        depreciation = 0
        if not cf.empty:
            year_cf = cf[cf['Năm'] == year].sort_values(by='Quý', ascending=False)
            if 'Khấu hao TSCĐ' in year_cf.columns:
                depreciation = year_cf.iloc[0]['Khấu hao TSCĐ'] if 'Khấu hao TSCĐ' in year_cf.columns else 0
            else:
                depreciation = 0
            
        year_data['income_statement']['ebitda'] = year_data['income_statement']['ebit'] + depreciation

        # Get cash flow data for the year
        if not cf.empty:
            year_cf = cf[cf['Năm'] == year].sort_values(by='Quý', ascending=False)
            if not year_cf.empty:
                last_quarter_cf = year_cf.iloc[0]
                
                # Safe data extraction with default values
                for field, cf_field in [
                    ('operating_cash_flow', 'Lưu chuyển tiền tệ ròng từ các hoạt động sản xuất kinh doanh (TT)'),
                    ('investing_cash_flow', 'Lưu chuyển tiền tệ ròng từ hoạt động đầu tư (TT)'),
                    ('financing_cash_flow', 'Lưu chuyển tiền tệ từ hoạt động tài chính (TT)')
                    ]:
                    if cf_field in last_quarter_cf and pd.notna(last_quarter_cf[cf_field]):
                        year_data['cash_flow'][field] = last_quarter_cf[cf_field]
        financial_data[str(int(year))] = year_data
        result['financial_data'] = financial_data
        result['years'] = [str(int(year)) for year in years]
        
        # Calculate financial ratios with proper error handling
        financial_ratios = {}
        
        for year in years:
            ratios = {
                # Profitability Ratios
                'ROA': 0,
                'ROE': 0,
                'ROS': 0,
                'Gross_Profit_Margin': 0,
                'EBIT_Margin': 0,
                'EBITDA_Margin': 0,
                
                # Liquidity Ratios
                'Current_Ratio': 0,
                'Quick_Ratio': 0,
                
                # Leverage Ratios
                'Debt_to_Equity': 0,
                'Debt_to_Assets': 0,
                'Interest_Coverage': 0,
                
                # Efficiency Ratios
                'Asset_Turnover': 0,
                'Inventory_Turnover': 0,
                'Receivables_Turnover': 0,
                'Working_Capital_Turnover': 0
            }
            year_str = str(int(year))
            year_data = financial_data.get(year_str, {})
            bs_data = year_data.get('balance_sheet', {})
            is_data = year_data.get('income_statement', {})
            
            # Calculate Profitability Ratios
            
            # ROA (Return on Assets)
            total_assets = bs_data.get('total_assets', 0)
            net_profit = is_data.get('net_profit', 0)
            if total_assets > 0 and net_profit != 0:
                ratios['ROA'] = (net_profit / total_assets) * 100
            
            # ROE (Return on Equity)
            equity = bs_data.get('equity', 0)
            if equity > 0 and net_profit != 0:
                ratios['ROE'] = (net_profit / equity) * 100

            # ROS (Return on Sales)
            revenue = is_data.get('revenue', 0)
            if revenue > 0 and net_profit != 0:
                ratios['ROS'] = (net_profit / revenue) * 100

            # Gross Profit Margin
            gross_profit = is_data.get('gross_profit', 0)
            if revenue > 0 and gross_profit != 0:
                ratios['Gross_Profit_Margin'] = (gross_profit / revenue) * 100

            # EBIT Margin
            ebit = is_data.get('ebit', 0)
            if revenue > 0 and ebit != 0:
                ratios['EBIT_Margin'] = (ebit / revenue) * 100

            # EBITDA Margin
            ebitda = is_data.get('ebitda', 0)
            if revenue > 0 and ebitda != 0:
                ratios['EBITDA_Margin'] = (ebitda / revenue) * 100

            # Calculate Liquidity Ratios

            # Current Ratio
            currrent_assets = bs_data.get('current_assets', 0)
            short_term_debt = bs_data.get('short_term_debt', 0)
            if short_term_debt > 0:
                ratios['Current_Ratio'] = current_assets / short_term_debt
            
            # Quick Ratio
            inventory = bs_data.get('inventory', 0)
            if short_term_debt > 0:
                ratios['Quick_Ratio'] = (current_assets - inventory) / short_term_debt
            
            # Calculate Leverage Ratios
            
            # Debt to Equity Ratio
            liabilities = bs_data.get('liabilities', 0)
            if equity > 0 and liabilities != 0:
                ratios['Debt_to_Equity'] = (liabilities / equity) * 100
            
            # Debt to Assets Ratio
            if total_assets > 0 and liabilities != 0:
                ratios['Debt_to_Assets'] = (liabilities / total_assets) * 100
            
            # Interest Coverage Ratio
            interest_expense = is_data.get('interest_expense', 0)
            if interest_expense > 0 and ebit != 0:
                ratios['Interest_Coverage'] = ebit / interest_expense
            
            # Calculate Efficiency Ratios
            
            # Asset Turnover
            if total_assets > 0 and revenue != 0:
                ratios['Asset_Turnover'] = revenue / total_assets
            
            # Inventory Turnover
            # Normally would use Cost of Goods Sold, but we can approximate using revenue if needed
            if inventory > 0 and revenue != 0:
                cogs = revenue - gross_profit if gross_profit > 0 else revenue * 0.7 # Approximate COGS if unavailable
                ratios['Inventory_Turnover'] = cogs / inventory
            
            # Receivables Turnover
            # Need to extract accounts receivable from balance sheet accounts_receivable = 0
            if not bs.empty:
                year_bs = data['balance_sheet'][data['balance_sheet']['Năm'] == year].sort_values(by='Quý', ascending=False)
                if not year_bs.empty and 'Các khoản phải thu ngắn hạn' in year_bs.iloc[0] and pd.notna( year_bs.iloc[0]['Các khoản phải thu ngắn hạn']):
                    accounts_receivable = year_bs.iloc[0]['Các khoản phải thu ngắn hạn']
            
            if accounts_receivable > 0 and revenue != 0:
                ratios['Receivables_Turnover'] = revenue / accounts_receivable
            
            # Working Capital Turnover
            working_capital = current_assets - short_term_debt
            if working_capital > 0 and revenue != 0:
                ratios['Working_Capital_Turnover'] = revenue / working_capital
            
            financial_ratios[year_str] = ratios
            
            result['financial_ratios'] = financial_ratios
            
            # Generate charts safely
            try:
                if len(years) > 0:
                    prepare_financial_charts(result)
                else:
                    # Initialize empty chart placeholders
                    result['revenue_profit_chart'] = None
                    result['ratios_chart'] = None
                    result['balance_sheet_chart'] = None
                    result['comparison_chart'] = None
            except Exception as e:
                print(f"Error preparing charts for {company_code}: {e}") 
                traceback.print_exc()
                result['revenue_profit_chart'] = None
                result['ratios_chart'] = None
                result['balance_sheet_chart'] = None
                result['comparison_chart'] = None
                # Ensure this statement is inside a function or remove it if unnecessary
                pass # Replace with appropriate logic if needed
            
            # ===== 1. Generate Financial Forecast =====
            try:
                result['financial_forecast'] = generate_financial_forecast(financial_data, years)
                result['forecast_years'] = [str(int(years[-1]) + i + 1) for i in range(3)] # Next 3 years
                
                # Generate forecast chart
                if result['financial_forecast']:
                    result['forecast_chart'] = generate_forecast_chart( 
                        years,
                        result['forecast_years'], 
                        financial_data,
                        result['financial_forecast']
                    )
                else:
                    result['forecast_chart'] = None
            except Exception as e:
                print(f"Error generating financial forecasts: {e}") 
                traceback.print_exc()
                result['financial_forecast'] = None
                result['forecast_years'] = None
                result['forecast_chart'] = None
            
            # ===== 2. Generate Investment Recommendation =====
            try:
                result['recommendation'] = generate_recommendation(
                    company_code,
                    financial_data,
                    financial_ratios,
                    years,
                    result.get('sector_avg', {})
                )
                pass
            except Exception as e:
                print(f"Error generating recommendation: {e}") 
                traceback.print_exc()
                result['recommendation'] = None
            except Exception as e:
                result['error'] = str(e) # Ghi lại lỗi vào result
                print(f"Error processing company {company_code}: {e}")
                return result

# New helper functions for financial forecasting and analysis
def generate_financial_forecast(financial_data, years):
    """
    Generate financial forecasts for the next 3 years based on historical data, using only financial statement data without market values
    """
    if not years or len(years) < 2:
        return None
    forecast_data = {}
    forecast_years = [int(years[-1]) + i + 1 for i in range(3)] # Next 3 years
    
    # Convert years to strings for dictionary access
    year_strs = [str(int(year)) for year in years]
    
    # Calculate growth rates from historical data
    growth_rates = {
        'revenue': 0,
        'gross_profit': 0,
        'net_profit': 0,
        'total_assets': 0,
        'equity': 0
    }
    
    # Use the last 3 years (or fewer if not available) to calculate average growth rates
    historical_years = year_strs[-3:] if len(year_strs) >= 3 else year_strs
    
    for metric in growth_rates.keys():
        values = []
        for i in range(1, len(historical_years)):
            prev_year = historical_years[i - 1]
            curr_year = historical_years[i]
            
            # Get values from income statement
            if metric in ['revenue', 'gross_profit', 'net_profit']:
                prev_value = financial_data.get(prev_year, {}).get('income_statement', {}).get(metric, 0)
                curr_value = financial_data.get(curr_year, {}).get('income_statement', {}).get(metric, 0)
            # Get values from balance sheetelse:
            else:
                prev_value = financial_data.get(prev_year, {}).get('balance_sheet', {}).get(metric, 0)
                curr_value = financial_data.get(curr_year, {}).get('balance_sheet', {}).get(metric, 0)
                
            # Calculate growth rate if both values are valid
            if prev_value > 0 and curr_value > 0:
                annual_growth = (curr_value / prev_value) - 1
                values.append(annual_growth)
        
        # Calculate average growth rate, use a moderate default if no valid data
        growth_rates[metric] = sum(values) / len(values) if values else 0.05
        
        # Cap extreme growth rates to reasonable values (-20% to 30%)
        growth_rates[metric] = max(min(growth_rates[metric], 0.3), -0.2)
        
    # Get latest financial values as base for forecasting
    latest_year = year_strs[-1]
    latest_data = financial_data.get(latest_year, {})

    latest_revenue = latest_data.get('income_statement', {}).get('revenue', 0)
    latest_gross_profit = latest_data.get('income_statement', {}).get('gross_profit', 0)
    latest_net_profit = latest_data.get('income_statement', {}).get('net_profit', 0)
    latest_total_assets = latest_data.get('balance_sheet', {}).get('total_assets', 0)
    latest_equity = latest_data.get('balance_sheet', {}).get('equity', 0)
    
    # Generate forecast for each year
    for i, year in enumerate(forecast_years):
        # Apply compound growth for each year in the forecast
        forecast_multiplier = (1 + growth_rates['revenue']) ** (i + 1)
        gross_profit_multiplier = (1 + growth_rates['gross_profit']) ** (i + 1)
        net_profit_multiplier = (1 + growth_rates['net_profit']) ** (i + 1)
        assets_multiplier = (1 + growth_rates['total_assets']) ** (i + 1)
        equity_multiplier = (1 + growth_rates['equity']) ** (i + 1)
        
        forecast_revenue = latest_revenue * forecast_multiplier
        forecast_gross_profit = latest_gross_profit * gross_profit_multiplier
        forecast_net_profit = latest_net_profit * net_profit_multiplier
        forecast_total_assets = latest_total_assets * assets_multiplier
        forecast_equity = latest_equity * equity_multiplier
        
        # Calculate EBIT (approximate as 1.3x net profit if not enough data)
        forecast_ebit = forecast_net_profit * 1.3
        
        # Calculate profit before tax (approximate as 1.2x net profit)
        forecast_profit_before_tax = forecast_net_profit * 1.2
        
        forecast_data[str(year)] = {
            'revenue': forecast_revenue,
            'gross_profit': forecast_gross_profit,
            'profit_before_tax': forecast_profit_before_tax,
            'net_profit': forecast_net_profit,
            'total_assets': forecast_total_assets,
            'equity': forecast_equity,
            'ebit': forecast_ebit
        }
    
    return forecast_data

def get_z_score_interpretation(z_score):
    """
    Interpret Altman's Z-Score for financial strength assessment
    """
    if z_score > 2.9:
        return "Vùng an toàn - Rủi ro phá sản thấp"
    elif z_score > 1.23:
        return "Vùng xám - Cần theo dõi"
    else:
        return "Vùng nguy hiểm - Rủi ro tài chính cao"

def generate_recommendation(company_code, financial_data, financial_ratios, years, sector_avg=None):
    """
    Tạo đánh giá tài chính đơn giản và các khuyến nghị chung
    """
    if not years:
        return None
    
    # Khởi tạo các thành phần khuyến nghị
    outlook = "Trung lập" # Triển vọng mặc định
    reasons = []
    conclusion = "" 
    
    # Lấy chỉ số tài chính mới nhất
    latest_year = str(int(years[-1]))
    latest_ratios = financial_ratios.get(latest_year, {})
    
    # Phân tích các chỉ số chính
    roe = latest_ratios.get('ROE', 0)
    roa = latest_ratios.get('ROA', 0)
    current_ratio = latest_ratios.get('Current_Ratio', 0)
    debt_to_equity = latest_ratios.get('Debt_to_Equity', 0)
    
    # Lấy chỉ số trung bình ngành nếu có
    sector_roe = sector_avg.get('Average ROE', 10) if sector_avg else 10
    sector_roa = sector_avg.get('Average ROA', 5) if sector_avg else 5
    sector_de = sector_avg.get('Average D/E Ratio', 100) if sector_avg else 100
    
    # Tính điểm sức khỏe tài chính cơ bản
    score = 0
    
    # Kiểm tra ROE
    if roe > sector_roe * 1.1:
        score += 1
        reasons.append(f"ROE ({roe:.2f}%) cao hơn trung bình ngành ({sector_roe:.2f}%)")
    elif roe < sector_roe * 0.9:
        score -= 1        
    
    # Kiểm tra ROA
    if roa > sector_roa * 1.1:
        score += 1
        reasons.append(f"ROA ({roa:.2f}%) cao hơn trung bình ngành ({sector_roa:.2f}%)")
    elif roa < sector_roa * 0.9:
        score -= 1
        
    # Kiểm tra Current Ratio
    if current_ratio > 1.5:
        score += 1
        reasons.append(f"Tỷ số thanh toán hiện hành tốt ({current_ratio:.2f})")
    elif current_ratio < 1.0:
        score -= 1
    
    # Kiểm tra Debt-to-Equity
    if debt_to_equity < sector_de * 0.8:
        score += 1
        reasons.append("Tỷ lệ nợ/vốn chủ sở hữu thấp, giảm rủi ro tài chính")
    elif debt_to_equity > sector_de * 1.2:
        score -= 1
        
    # Kiểm tra xu hướng tăng trưởng nếu có đủ dữ liệu lịch sử
    if len(years) >= 3:
        # Tính tăng trưởng doanh thu
        recent_years = [str(int(y)) for y in years[-3:]]
        revenue_values = [financial_data.get(y, {}).get('income_statement', {}).get('revenue', 0) for y in recent_years]
        profit_values = [financial_data.get(y, {}).get('income_statement', {}).get('net_profit', 0) for y in recent_years]
        
        # Kiểm tra nếu có dữ liệu hợp lệ
        if all(v > 0 for v in revenue_values) and all(v > 0 for v in profit_values):
            # Tính tỷ lệ tăng trưởng trung bình
            revenue_growth = [(revenue_values[i] / revenue_values[i - 1] - 1) * 100 for i in range(1, len(revenue_values))]
            profit_growth = [(profit_values[i] / profit_values[i - 1] - 1) * 100 for i in range(1, len(profit_values))]

            avg_revenue_growth = sum(revenue_growth) / len(revenue_growth)
            avg_profit_growth = sum(profit_growth) / len(profit_growth)
            
            # Đánh giá tăng trưởng
            if avg_revenue_growth > 10 and avg_profit_growth > 10:
                score += 1
                reasons.append(f"Tăng trưởng mạnh về doanh thu({avg_revenue_growth:.2f}%) và lợi nhuận ({avg_profit_growth:.2f}%)")
            elif avg_revenue_growth < 0 and avg_profit_growth < 0:
                score -= 1
                reasons.append(f"Doanh thu ({avg_revenue_growth:.2f}%) và lợi nhuận({avg_profit_growth:.2f}%) suy giảm")
            elif avg_revenue_growth > 5 or avg_profit_growth > 5:
                reasons.append(f"Tăng trưởng ổn định về doanh thu và lợi nhuận")
    
    # Xác định triển vọng dựa trên điểm số
    if score >= 3:
        outlook = "Tích cực cao"
    elif score >= 1:
        outlook = "Tích cực"
    elif score <= -3:
        outlook = "Tiêu cực"
    elif score <= -1:
        outlook = "Thận trọng"
    else:
        outlook = "Trung lập"
    
    # Tạo kết luận dựa trên triển vọng
    if outlook == "Tích cực cao":
        conclusion = f"Công ty {company_code} có tình hình tài chính rất tốt với khả năng sinh lời cao và tăng trưởng mạnh. Triển vọng phát triển trong tương lai rất tích cực."
    elif outlook == "Tích cực":
        conclusion = f"Công ty {company_code} có tình hình tài chính tốt với các chỉ số tài chính vượt trội so với trung bình ngành. Triển vọng phát triển trong tương lai tích cực."
    elif outlook == "Tiêu cực":
        conclusion = f"Công ty {company_code} đang gặp nhiều khó khăn trong hoạt động kinh doanh, thể hiện qua các chỉ số tài chính kém. Cần cải thiện đáng kể để phát triển bền vững."
    elif outlook == "Thận trọng":
        conclusion = f"Công ty {company_code} có một số điểm yếu về tài chính cần được cải thiện. Triển vọng phát triển trong tương lai cần được theo dõi thận trọng."
    else:
        conclusion = f"Công ty {company_code} có tình hình tài chính ở mức trung bình. Triển vọng phát triển trong tương lai phụ thuộc vào khả năng cải thiện các chỉ số tài chính."
    
    # Đảm bảo có ít nhất 3 lý do
    if len(reasons) < 3:
        default_reasons = [
            "Chỉ số tài chính ở mức trung bình ngành",
            "Cơ cấu tài chính tương đối cân đối",
            "Tốc độ tăng trưởng ổn định"
        ]
        reasons.extend(default_reasons[:3 - len(reasons)])
    
    return {
        'outlook': outlook,
        'reasons': reasons,
        'conclusion': conclusion
    }

def generate_forecast_chart(historical_years, forecast_years, historical_data, forecast_data):
    """
    Generate an improved chart for financial forecasts with better visual clarity
    """
    try:
        # Set a clean, professional style
        plt.style.use('seaborn-v0_8-whitegrid')
        
        # Create a larger figure with better proportions
        plt.figure(figsize=(12, 7))
        
        # Ensure proper DPI for clear rendering
        plt.rcParams['figure.dpi'] = 100
        
        # Convert years to strings for consistent handling
        historical_years = [str(int(year)) for year in historical_years]
        
        # Extract historical revenue and profit data
        historical_revenue = [historical_data.get(year, {}).get('income_statement', {}).get('revenue', 0) / 1e9 for year in historical_years]
        historical_profit = [historical_data.get(year, {}).get('income_statement', {}).get('net_profit', 0) / 1e9 for year in historical_years]
        
        # Extract forecast revenue and profit data
        forecast_revenue = [forecast_data.get(year, {}).get('revenue', 0) / 1e9 for year in forecast_years]
        forecast_profit = [forecast_data.get(year, {}).get('net_profit', 0) / 1e9 for year in forecast_years]
        
        # Skip if all values are zero
        if sum(historical_revenue) == 0 and sum(historical_profit) == 0 and sum(forecast_revenue) == 0 and sum(forecast_profit) == 0:
            return None
        
        # Combine all years for the x-axis
        all_years = historical_years + forecast_years
        
        # Plot historical data with solid lines and clear markers
        plt.plot(historical_years, historical_revenue, marker='o', linestyle='-', linewidth=2.5, color='#1f77b4', label='Doanh thu thực tế')
        plt.plot(historical_years, historical_profit, marker='o', linestyle='-', linewidth=2.5, color='#2ca02c', label='Lợi nhuận thực tế')
        
        # Plot forecast data with dashed lines and distinct markers
        plt.plot(forecast_years, forecast_revenue, marker='s', linestyle='--', linewidth=2.5, color='#1f77b4', label='Doanh thu dự báo')
        plt.plot(forecast_years, forecast_profit, marker='s', linestyle='--', linewidth=2.5, color='#2ca02c', label='Lợi nhuận dự báo')
        
        # Improve grid appearance
        plt.grid(True, linestyle='--', alpha=0.7)
        
        # Add a dividing line between historical and forecast data
        plt.axvline(x=historical_years[-1], color='gray', linestyle='--', alpha=0.7)
        
        # Add "Forecast" annotation at the dividing line
        mid_y = (max(historical_revenue + forecast_revenue) + min(historical_profit + forecast_profit)) / 2
        plt.text(historical_years[-1], mid_y, 'Dự báo', ha='right', va='center', bbox=dict(facecolor='#fff9cc', alpha=0.8, edgecolor='#e6e6e6', boxstyle='round,pad=0.5'))
        
        # Add value labels for key points (first, last historical, and ast forecast)
        #  Revenue labels
        plt.text(historical_years[0], historical_revenue[0], f'{historical_revenue[0]:.1f}', ha='center', va='bottom', fontsize=9)
        plt.text(historical_years[-1], historical_revenue[-1], f'{historical_revenue[-1]:.1f}', ha='center', va='bottom', fontsize=9)
        plt.text(forecast_years[-1], forecast_revenue[-1], f'{forecast_revenue[-1]:.1f}', ha='center', va='bottom', fontsize=9)

        # Profit labels
        plt.text(historical_years[0], historical_profit[0], f'{historical_profit[0]:.1f}', ha='center', va='top', fontsize=9)
        plt.text(historical_years[-1], historical_profit[-1], f'{historical_profit[-1]:.1f}', ha='center', va='top', fontsize=9)
        plt.text(forecast_years[-1], forecast_profit[-1], f'{forecast_profit[-1]:.1f}', ha='center', va='top', fontsize=9)
        
        # Improve axis labels and title
        plt.xlabel('Năm', fontsize=12, fontweight='bold')
        plt.ylabel('Tỷ VNĐ', fontsize=12, fontweight='bold')
        plt.title('Dự báo kết quả kinh doanh', fontsize=16, fontweight='bold', pad=20)
        
        # Set tick parameters for better readability
        plt.tick_params(axis='both', which='major', labelsize=10)
        
        # Improve legend appearance and placement
        plt.legend(loc='best', frameon=True, fancybox=True, shadow=True, fontsize=10)
        
        # Tighten layout to use space efficiently
        plt.tight_layout()
        
        # Add subtle background color to differentiate forecast area
        ax = plt.gca()
        forecast_idx = len(historical_years)
        xlim = ax.get_xlim()
        ylim = ax.get_ylim()
        
        # Add subtle background shading for the forecast region
        rect = plt.Rectangle((forecast_idx - 0.5, ylim[0]), xlim[1] - forecast_idx + 0.5, ylim[1] - ylim[0], color='#f9f9f9', alpha=0.3, zorder=-1)
        ax.add_patch(rect)
        
        # Save to buffer
        buf = io.BytesIO()
        plt.savefig(buf, format='png', bbox_inches='tight', dpi=100)
        buf.seek(0)
        
        # Convert to base64
        forecast_chart = base64.b64encode(buf.getvalue()).decode('utf-8')
        plt.close()
        
        return forecast_chart
    
    except Exception as e:
        print(f"Error generating forecast chart: {e}")
        traceback.print_exc()
        return None

def generate_valuation_chart(valuation_data, sector_avg):
    """
    Generate chart comparing valuation metrics with sector averages
    """
    try:
        plt.figure(figsize=(10, 6))
        
        # Extract valuation metrics
        company_metrics = [
            valuation_data.get('PE', 0),
            valuation_data.get('PB', 0),
            valuation_data.get('PS', 0),
            valuation_data.get('EV_EBITDA', 0),
            valuation_data.get('Dividend_Yield', 0)
        ]
        
        # Extract sector average metrics with defaults
        sector_metrics = [
            sector_avg.get('Average PE', 15),
            sector_avg.get('Average PB', 1.5),
            sector_avg.get('Average PS', 1.5),
            sector_avg.get('Average EV_EBITDA', 10),
            sector_avg.get('Average Dividend_Yield', 3)
        ]
        
        # Skip if all values are zero
        if sum(company_metrics) == 0 and sum(sector_metrics) == 0:
            return None
        
        # Metrics labels
        labels = ['P/E', 'P/B', 'P/S', 'EV/EBITDA', 'Tỷ suất cổ tức (%)']
        
        # Set up bar positions
        x = np.arange(len(labels))
        width = 0.35
        
        # Create bars
        plt.bar(x - width / 2, company_metrics, width, label='Công ty', color='#3498db')
        plt.bar(x + width / 2, sector_metrics, width, label='Trung bình ngành', color='#e74c3c')
        
        # Add details
        plt.xlabel('Chỉ số định giá', fontsize=12)
        plt.ylabel('Giá trị', fontsize=12)
        plt.title('So sánh định giá với trung bình ngành', fontsize=14, fontweight='bold')
        plt.xticks(x, labels, fontsize=10)
        plt.legend(fontsize=10)
        plt.grid(True, linestyle='--', alpha=0.7)
        
        # Add value labels on bars
        for i, v in enumerate(company_metrics):
            if v > 0:
                plt.text(i - width / 2, v, f'{v:.2f}', ha='center', va='bottom', fontsize=9)
        
        for i, v in enumerate(sector_metrics):
            if v > 0:
                plt.text(i + width / 2, v, f'{v:.2f}', ha='center', va='bottom', fontsize=9)
                
        # Save to buffer
        buf = io.BytesIO()
        plt.savefig(buf, format='png', bbox_inches='tight')
        buf.seek(0)
        
        # Convert to base64
        valuation_chart = base64.b64encode(buf.getvalue()).decode('utf-8')
        plt.close()
        
        return valuation_chart
    
    except Exception as e:
        print(f"Error generating valuation chart: {e}")
        traceback.print_exc()
        return None

# Hàm tạo biểu đồ cho báo cáo
def prepare_financial_charts(data):
    years = data.get('years', [])
    if not years or len(years) < 2:
        # Need at least 2 years of data for meaningful charts
        data['revenue_profit_chart'] = None
        data['ratios_chart'] = None
        data['balance_sheet_chart'] = None
        data['comparison_chart'] = None
        # Initialize the new chart variables\
        data['profitability_chart'] = None
        data['growth_chart'] = None
        data['liquidity_chart'] = None
        data['leverage_chart'] = None
        data['efficiency_chart'] = None
        return
    
    # Thiết lập style cho biểu đồ
    plt.style.use('seaborn-v0_8-whitegrid')
    
    # Cấu hình font
    plt.rcParams['font.family'] = 'sans-serif'
    plt.rcParams['font.sans-serif'] = ['Arial', 'Helvetica', 'DejaVu Sans']
    
    # Đặt chất lượng biểu đồplt.rcParams['figure.dpi'] = 100
    plt.rcParams['savefig.dpi'] = 100
    plt.rcParams['savefig.dpi'] = 100
    
    # Generate original charts (keep existing code)
    # ...
    # Generate the new specialized charts
    try:
        generate_profitability_chart(data)
    except Exception as e:
        print(f"Error creating profitability chart: {e}")
        traceback.print_exc()
        data['profitability_chart'] = None
    
    def generate_growth_chart(data):
        """Generate chart for growth ratios comparison with sector average"""
        if 'company_metrics' not in data or 'sector_avg' not in data:
            data['growth_chart'] = None
            return
        
        plt.figure(figsize=(10, 6))
        
        # Extract growth metrics
        metrics = ['Revenue Growth (%)', 'Net Income Growth (%)', 'Total Assets Growth (%)']
        company_values = []
        sector_values = []
        
        # Get company metrics
        company_metrics = data['company_metrics']
        company_values = [
            company_metrics.get('Revenue Growth (%)', 0),
            company_metrics.get('Net Income Growth (%)', 0),
            company_metrics.get('Total Assets Growth (%)', 0)
        ]
        
        # Get sector averages
        sector_avg = data['sector_avg']
        sector_values = [
            sector_avg.get('Average Revenue Growth', 0),
            sector_avg.get('Average Net Income Growth', 0),
            sector_avg.get('Average Total Assets Growth', 0)
        ]
        
        # Skip if all values are zero
        if sum(company_values) == 0 and sum(sector_values) == 0:
            data['growth_chart'] = None
            return
        
        # Create the chart
        x = np.arange(len(metrics))
        width = 0.35
        plt.bar(x - width / 2, company_values, width, label=data['company_code'], color='#3498db')
        plt.bar(x + width / 2, sector_values, width, label='Trung bình ngành', color='#e74c3c')
        plt.xlabel('Chỉ số', fontsize=12)
        plt.ylabel('Phần trăm (%)', fontsize=12)
        plt.title('So sánh chỉ số tăng trưởng với trung bình ngành', fontsize=14, fontweight='bold')
        plt.xticks(x, metrics, fontsize=10)
        plt.legend(fontsize=10)
        plt.grid(True, linestyle='--', alpha=0.7)
        plt.tight_layout()
        
        # Add value labels
        for i, v in enumerate(company_values):
            plt.text(i - width / 2, v + (1 if v >= 0 else -3),f'{v:.1f}%', ha='center', fontsize=9)
        
        for i, v in enumerate(sector_values):
            plt.text(i + width / 2, v + (1 if v >= 0 else -3),f'{v:.1f}%', ha='center', fontsize=9)
        
        # Save chart
        buf = io.BytesIO()
        plt.savefig(buf, format='png', bbox_inches='tight')
        buf.seek(0)
        data['growth_chart'] = base64.b64encode(buf.getvalue()).decode('utf-8')
        plt.close()
    try:
        generate_growth_chart(data)
    except Exception as e:
        print(f"Error creating growth chart: {e}")
        traceback.print_exc()
        data['growth_chart'] = None
        traceback.print_exc()
        data['growth_chart'] = None

    def generate_liquidity_chart(data):
        """Generate chart for liquidity ratios comparison with sector average"""
        if 'company_metrics' not in data or 'sector_avg' not in data:
            data['liquidity_chart'] = None
            return
        
        plt.figure(figsize=(10, 6))
        
        # Extract liquidity metrics
        metrics = ['Current Ratio', 'Quick Ratio', 'Interest Coverage Ratio']
        company_values = [
            data['company_metrics'].get('Current Ratio', 0),
            data['company_metrics'].get('Quick Ratio', 0),
            data['company_metrics'].get('Interest Coverage Ratio', 0)
        ]
        sector_values = [
            data['sector_avg'].get('Average Current Ratio', 0),
            data['sector_avg'].get('Average Quick Ratio', 0),
            data['sector_avg'].get('Average Interest Coverage Ratio', 0)
        ]
        
        # Skip if all values are zero
        if sum(company_values) == 0 and sum(sector_values) == 0:
            data['liquidity_chart'] = None
            return
        
        # Create the chart
        x = np.arange(len(metrics))
        width = 0.35
        plt.bar(x - width / 2, company_values, width, label='Company', color='#3498db')
        plt.bar(x + width / 2, sector_values, width, label='Sector Average', color='#e74c3c')
        plt.xlabel('Metrics', fontsize=12)
        plt.ylabel('Values', fontsize=12)
        plt.title('Liquidity Ratios Comparison', fontsize=14, fontweight='bold')
        plt.xticks(x, metrics, fontsize=10)
        plt.legend(fontsize=10)
        plt.grid(True, linestyle='--', alpha=0.7)
        plt.tight_layout()
        
        # Add value labels
        for i, v in enumerate(company_values):
            if v > 0:
                plt.text(i - width / 2, v + 0.1, f'{v:.2f}', ha='center', fontsize=9)
        
        for i, v in enumerate(sector_values):
            if v > 0:
                plt.text(i + width / 2, v + 0.1, f'{v:.2f}', ha='center', fontsize=9)
        
        # Save chart
        buf = io.BytesIO()
        plt.savefig(buf, format='png', bbox_inches='tight')
        buf.seek(0)
        data['liquidity_chart'] = base64.b64encode(buf.getvalue()).decode('utf-8')
        plt.close()
    try:
        generate_liquidity_chart(data)
    except Exception as e:
        print(f"Error creating liquidity chart: {e}")
        traceback.print_exc()
        data['liquidity_chart'] = None
    
    def generate_leverage_chart(data):
        """Generate chart for leverage ratios comparison with sector average"""
        if 'company_metrics' not in data or 'sector_avg' not in data:
            data['leverage_chart'] = None
            return
            
        plt.figure(figsize=(10, 6))
        
        # Extract leverage metrics
        metrics = ['Debt to Equity (%)', 'Debt to Assets (%)', 'Interest Coverage']
        company_values = [
            data['company_metrics'].get('Debt_to_Equity', 0),
            data['company_metrics'].get('Debt_to_Assets', 0),
            data['company_metrics'].get('Interest_Coverage', 0)
        ]
        sector_values = [
            data['sector_avg'].get('Average Debt to Equity', 0),
            data['sector_avg'].get('Average Debt to Assets', 0),
            data['sector_avg'].get('Average Interest Coverage', 0)
        ]
        
        # Skip if all values are zero
        if sum(company_values) == 0 and sum(sector_values) == 0:
            data['leverage_chart'] = None
            return
        
        # Create the chart
        x = np.arange(len(metrics))
        width = 0.35
        plt.bar(x - width / 2, company_values, width, label='Company', color='#3498db')
        plt.bar(x + width / 2, sector_values, width, label='Sector Average', color='#e74c3c')
        plt.xlabel('Metrics', fontsize=12)
        plt.ylabel('Values', fontsize=12)
        plt.title('Leverage Ratios Comparison', fontsize=14, fontweight='bold')
        plt.xticks(x, metrics, fontsize=10)
        plt.legend(fontsize=10)
        plt.grid(True, linestyle='--', alpha=0.7)
        plt.tight_layout()
        
        # Add value labels
        for i, v in enumerate(company_values):
            if v > 0:
                plt.text(i - width / 2, v + 0.1, f'{v:.2f}', ha='center', fontsize=9)
        
        for i, v in enumerate(sector_values):
            if v > 0:
                plt.text(i + width / 2, v + 0.1, f'{v:.2f}', ha='center', fontsize=9) 
        
        # Save chart
        buf = io.BytesIO()
        plt.savefig(buf, format='png', bbox_inches='tight')
        buf.seek(0)
        data['leverage_chart'] = base64.b64encode(buf.getvalue()).decode('utf-8')
        plt.close()
    try:
        generate_leverage_chart(data)
    except Exception as e:
        print(f"Error creating leverage chart: {e}")
        traceback.print_exc()
        data['leverage_chart'] = None
    
    def generate_efficiency_chart(data):
        """Generate chart for efficiency ratios comparison with sector average"""
        if 'company_metrics' not in data or 'sector_avg' not in data:
            data['efficiency_chart'] = None
            return
        
        plt.figure(figsize=(10, 6))
        
        # Extract efficiency metrics
        metrics = ['Asset Turnover', 'Inventory Turnover', 'Receivables Turnover', 'Working Capital Turnover']
        company_values = [
            data['company_metrics'].get('Asset_Turnover', 0),
            data['company_metrics'].get('Inventory_Turnover', 0),
            data['company_metrics'].get('Receivables_Turnover', 0),
            data['company_metrics'].get('Working_Capital_Turnover', 0)
        ]
        sector_values = [
            data['sector_avg'].get('Average Asset Turnover', 0),
            data['sector_avg'].get('Average Inventory Turnover', 0),
            data['sector_avg'].get('Average Receivables Turnover', 0),data['sector_avg'].get('Average Working Capital Turnover', 0)
        ]
        
        # Skip if all values are zero
        if sum(company_values) == 0 and sum(sector_values) == 0:
            data['efficiency_chart'] = None
            return
        
        # Create the chart
        x = np.arange(len(metrics))
        width = 0.35
        plt.bar(x - width / 2, company_values, width, label='Company', color='#3498db')
        plt.bar(x + width / 2, sector_values, width, label='Sector Average', color='#e74c3c')
        plt.xlabel('Metrics', fontsize=12)
        plt.ylabel('Values', fontsize=12)
        plt.title('Efficiency Ratios Comparison', fontsize=14, fontweight='bold')
        plt.xticks(x, metrics, fontsize=10, rotation=45, ha='right')
        plt.legend(fontsize=10)
        plt.grid(True, linestyle='--', alpha=0.7)
        plt.tight_layout()
        
        # Add value labels
        for i, v in enumerate(company_values):
            if v > 0:
                plt.text(i - width / 2, v + 0.1, f'{v:.2f}', ha='center', fontsize=9)
        
        for i, v in enumerate(sector_values):
            if v > 0:
                plt.text(i + width / 2, v + 0.1, f'{v:.2f}', ha='center', fontsize=9)
                
        # Save chart
        buf = io.BytesIO()
        plt.savefig(buf, format='png', bbox_inches='tight')
        buf.seek(0)
        data['efficiency_chart'] = base64.b64encode(buf.getvalue()).decode('utf-8')
        plt.close()
        """Generate chart for efficiency ratios comparison with sector average"""
        if 'company_metrics' not in data or 'sector_avg' not in data:
            data['efficiency_chart'] = None
            return
        
        plt.figure(figsize=(10, 6))
        # Extract efficiency metrics
        metrics = ['Asset Turnover', 'Inventory Turnover', 'Receivables Turnover', 'Working Capital Turnover']
        company_values = [
            data['company_metrics'].get('Asset_Turnover', 0),
            data['company_metrics'].get('Inventory_Turnover', 0),
            data['company_metrics'].get('Receivables_Turnover', 0),data['company_metrics'].get('Working_Capital_Turnover', 0)
        ]
        sector_values = [
            data['sector_avg'].get('Average Asset Turnover', 0),
            data['sector_avg'].get('Average Inventory Turnover', 0),
            data['sector_avg'].get('Average Receivables Turnover', 0),
            data['sector_avg'].get('Average Working Capital Turnover', 0)
        ]
        
        # Skip if all values are zero
        if sum(company_values) == 0 and sum(sector_values) == 0:
            data['efficiency_chart'] = None
            return
        
        # Create the chart
        x = np.arange(len(metrics))
        width = 0.35
        plt.bar(x - width / 2, company_values, width, label='Company', color='#3498db')
        plt.bar(x + width / 2, sector_values, width, label='Sector Average', color='#e74c3c')
        plt.xlabel('Metrics', fontsize=12)
        plt.ylabel('Values', fontsize=12)
        plt.title('Efficiency Ratios Comparison', fontsize=14, fontweight='bold')
        plt.xticks(x, metrics, fontsize=10, rotation=45, ha='right')
        plt.legend(fontsize=10)
        plt.grid(True, linestyle='--', alpha=0.7)
        plt.tight_layout()
        
        # Add value labels
        for i, v in enumerate(company_values):
            if v > 0:
                plt.text(i - width / 2, v + 0.1, f'{v:.2f}', ha='center', fontsize=9)
        
        for i, v in enumerate(sector_values):
            if v > 0:
                plt.text(i + width / 2, v + 0.1, f'{v:.2f}', ha='center', fontsize=9)
        
        # Save chart
        buf = io.BytesIO()
        plt.savefig(buf, format='png', bbox_inches='tight')
        buf.seek(0)
        data['efficiency_chart'] = base64.b64encode(buf.getvalue()).decode('utf-8')
        plt.close()
    try:
        generate_efficiency_chart(data)
    except Exception as e:
        print(f"Error creating efficiency chart: {e}")
        traceback.print_exc()
        data['efficiency_chart'] = None
        
def generate_profitability_chart(data):
    """Generate chart for profitability ratios comparison with sector average"""
    if 'company_metrics' not in data or 'sector_avg' not in data:
          data['profitability_chart'] = None
          return
    plt.figure(figsize=(10, 6))
    
    # Extract profitability metrics
    metrics = ['ROA (%)', 'ROE (%)', 'ROS (%)', 'EBIT Margin (%)','EBITDA Margin (%)', 'Gross Profit Margin (%)']
    company_values = []
    sector_values = []
    
    # Get company metrics
    company_metrics = data['company_metrics']
    company_values = [
        company_metrics.get('ROA (%)', 0),
        company_metrics.get('ROE (%)', 0),
        company_metrics.get('ROS (%)', 0),
        company_metrics.get('EBIT Margin (%)', 0),
        company_metrics.get('EBITDA Margin (%)', 0),
        company_metrics.get('Gross Profit Margin (%)', 0)
    ]
    
    # Get sector averages
    sector_avg = data['sector_avg']
    sector_values = [
        sector_avg.get('Average ROA', 0),
        sector_avg.get('Average ROE', 0),
        sector_avg.get('Average ROS', 0),
        sector_avg.get('Average EBIT Margin', 0),
        sector_avg.get('Average EBITDA Margin', 0),
        sector_avg.get('Average Gross Profit Margin', 0)
    ]
        
    # Skip if all values are zero
    if sum(company_values) == 0 and sum(sector_values) == 0:
        data['profitability_chart'] = None
        return
    
    # Create the chart
    x = np.arange(len(metrics))
    width = 0.35
    plt.bar(x - width / 2, company_values, width, label=data['company_code'], color='#3498db')
    plt.bar(x + width / 2, sector_values, width, label='Trung bình ngành', color='#e74c3c')
    plt.xlabel('Chỉ số', fontsize=12)
    plt.ylabel('Phần trăm (%)', fontsize=12)
    plt.title('So sánh chỉ số sinh lời với trung bình ngành', fontsize=14, fontweight='bold')
    plt.xticks(x, metrics, fontsize=10, rotation=45, ha='right')
    plt.legend(fontsize=10)
    plt.grid(True, linestyle='--', alpha=0.7)
    plt.tight_layout()
        
    # Add value labels
    for i, v in enumerate(company_values):
        if v > 0:
            plt.text(i - width / 2, v + 0.5, f'{v:.1f}%', ha='center', fontsize=9)
        
    for i, v in enumerate(sector_values):
        if v > 0:
            plt.text(i + width / 2, v + 0.5, f'{v:.1f}%', ha='center', fontsize=9)
        
    # Save chart
    buf = io.BytesIO()
    plt.savefig(buf, format='png', bbox_inches='tight')
    buf.seek(0)
    data['profitability_chart'] = base64.b64encode(buf.getvalue()).decode('utf-8')
    plt.close()
        
    def generate_growth_chart(data):
        """Generate chart for growth ratios comparison with sector average"""
        if 'company_metrics' not in data or 'sector_avg' not in data:
            data['growth_chart'] = None
            return
        
        plt.figure(figsize=(10, 6))
        
        # Extract growth metrics
        metrics = ['Revenue Growth (%)', 'Net Income Growth (%)', 'Total Assets Growth (%)']
        company_values = []
        sector_values = []
        
        # Get company metrics
        company_metrics = data['company_metrics']
        company_values = [
            company_metrics.get('Revenue Growth (%)', 0),
            company_metrics.get('Net Income Growth (%)', 0),
            company_metrics.get('Total Assets Growth (%)', 0)]
        
        # Get sector averages
        Sector_avg = data['sector_avg']
        sector_values = [
            sector_avg.get('Average Revenue Growth', 0),
            sector_avg.get('Average Net Income Growth', 0),
            sector_avg.get('Average Total Assets Growth', 0)
        ]
        
        # Skip if all values are zero
        if sum(company_values) == 0 and sum(sector_values) == 0:
            data['growth_chart'] = None
            return
        
        # Create the chart
        x = np.arange(len(metrics))
        width = 0.35
        plt.bar(x - width / 2, company_values, width, label=data['company_code'], color='#3498db')
        plt.bar(x + width / 2, sector_values, width, label='Trung bình ngành', color='#e74c3c')
        plt.xlabel('Chỉ số', fontsize=12)
        plt.ylabel('Phần trăm (%)', fontsize=12)
        plt.title('So sánh chỉ số tăng trưởng với trung bình ngành', fontsize=14, fontweight='bold')
        plt.xticks(x, metrics, fontsize=10)
        plt.legend(fontsize=10)
        plt.grid(True, linestyle='--', alpha=0.7)
        plt.tight_layout()
        
        # Add value labels
        for i, v in enumerate(company_values):
            plt.text(i - width / 2, v + (1 if v >= 0 else -3),f'{v:.1f}%', ha='center', fontsize=9)
        
        for i, v in enumerate(sector_values):
            plt.text(i + width / 2, v + (1 if v >= 0 else -3),f'{v:.1f}%', ha='center', fontsize=9)
        
        # Save chart
        buf = io.BytesIO()
        plt.savefig(buf, format='png', bbox_inches='tight')
        buf.seek(0)
        data['growth_chart'] = base64.b64encode(buf.getvalue()).decode('utf-8')
        plt.close()
        
    def generate_liquidity_chart(data): 
        """Generate chart for liquidity ratios comparison with sector average"""
        if 'company_metrics' not in data or 'sector_avg' not in data:data['liquidity_chart'] = None
        return
    
    plt.figure(figsize=(10, 6))
    
    # Extract liquidity metrics
    metrics = ['Current Ratio', 'Quick Ratio', 'Interest Coverage Ratio']
    company_values = []
    sector_values = []
    
    # Get company metrics
    company_metrics = data['company_metrics']
    company_values = [
        company_metrics.get('Current Ratio', 0),
        company_metrics.get('Quick Ratio', 0),
        company_metrics.get('Interest Coverage Ratio', 0)
    ]
    
    # Get sector averages
    sector_avg = data['sector_avg']
    sector_values = [
        sector_avg.get('Average Current Ratio', 0),
        sector_avg.get('Average Quick Ratio', 0),
        sector_avg.get('Average Interest Coverage Ratio', 0)
    ]
    
    # Skip if all values are zero
    if sum(company_values) == 0 and sum(sector_values) == 0:
        data['liquidity_chart'] = None
        return
    
    # Create the chart
    x = np.arange(len(metrics))
    width = 0.35
    plt.bar(x - width / 2, company_values, width, label=data['company_code'], color='#3498db')
    plt.bar(x + width / 2, sector_values, width, label='Trung bình ngành', color='#e74c3c')
    plt.xlabel('Chỉ số', fontsize=12)
    plt.ylabel('Lần', fontsize=12)
    plt.title('So sánh chỉ số thanh khoản với trung bình ngành', fontsize=14, fontweight='bold')
    plt.xticks(x, metrics, fontsize=10)
    plt.legend(fontsize=10)
    plt.grid(True, linestyle='--', alpha=0.7)
    plt.tight_layout()
    
    # Add value labels
    for i, v in enumerate(company_values):
        if v > 0:plt.text(i - width / 2, v + 0.1, f'{v:.2f}', ha='center', fontsize=9)
    
    for i, v in enumerate(sector_values):
        if v > 0:
            plt.text(i + width / 2, v + 0.1, f'{v:.2f}', ha='center', fontsize=9)
    
    # Save chart
    buf = io.BytesIO()
    plt.savefig(buf, format='png', bbox_inches='tight')
    buf.seek(0)
    data['liquidity_chart'] = base64.b64encode(buf.getvalue()).decode('utf-8')
    plt.close()

def get_financial_statements(company_code):
    """API endpoint for retrieving financial statements by year"""
    year = request.args.get('year', None)
    
    if not year:
        return jsonify({"error": "Year parameter is required"}), 400
    
    try:
        year = int(year)
    except ValueError:
        return jsonify({"error": "Year must be a valid integer"}), 400
    
    # Initialize result structure
    result = {
        "balance_sheet": None,
        "income_statement": None,
        "cash_flow": None
    }
    
    try:
        # Get balance sheet data
        if 'balance_sheet' in data:
            balance_sheet_data = data['balance_sheet'][
                (data['balance_sheet']['Mã'] == company_code) &
                (data['balance_sheet']['Năm'] == year)
                ]
                
            if not balance_sheet_data.empty:
                # Sort by quarter to get the latest quarter data
                balance_sheet_data = balance_sheet_data.sort_values(by='Quý', ascending=False)
                latest_bs = balance_sheet_data.iloc[0]
                
                # Extract relevant fields
                result["balance_sheet"] = {
                    "current_assets": float(
                        latest_bs['TÀI SẢN NGẮN HẠN']) if 'TÀI SẢN NGẮN HẠN' in latest_bs and pd.notna(
                        latest_bs['TÀI SẢN NGẮN HẠN']) else None, "cash_and_equivalents": float(
                        latest_bs['Tiền và tương đương tiền']) if 'Tiền và tương đương tiền' in latest_bs and pd.notna(
                        latest_bs['Tiền và tương đương tiền']) else None, "short_term_investments": float(latest_bs['Đầu tư tài chính ngắn hạn']) if 'Đầu tư tài chính ngắn hạn' in latest_bs and pd.notna(
                        latest_bs['Đầu tư tài chính ngắn hạn']) else None, "short_term_receivables": float(latest_bs['Các khoản phải thu ngắn hạn']) if 'Các khoản phải thu ngắn hạn' in latest_bs and pd.notna(
                        latest_bs['Các khoản phải thu ngắn hạn']) else None, "inventory": float(
                        latest_bs['Hàng tồn kho, ròng']) if 'Hàng tồn kho, ròng' in latest_bs and pd.notna(
                        latest_bs['Hàng tồn kho, ròng']) else None, "other_current_assets": float(
                        latest_bs['Tài sản ngắn hạn khác']) if 'Tài sản ngắn hạn khác' in latest_bs and pd.notna(
                        latest_bs['Tài sản ngắn hạn khác']) else None, "non_current_assets": float(
                        latest_bs['TÀI SẢN DÀI HẠN']) if 'TÀI SẢN DÀI HẠN' in latest_bs and pd.notna(
                        latest_bs['TÀI SẢN DÀI HẠN']) else None, "fixed_assets": float(latest_bs['Tài sản cố định']) if 'Tài sản cố định' in latest_bs and pd.notna(
                        latest_bs['Tài sản cố định']) else None, "long_term_investments": float(
                        latest_bs['Đầu tư dài hạn']) if 'Đầu tư dài hạn' in latest_bs and pd.notna(
                        latest_bs['Đầu tư dài hạn']) else None, "other_non_current_assets": float(
                        latest_bs['Tài sản dài hạn khác']) if 'Tài sản dài hạn khác' in latest_bs and pd.notna(
                        latest_bs['Tài sản dài hạn khác']) else None, "total_assets": float(
                        latest_bs['TỔNG CỘNG TÀI SẢN']) if 'TỔNG CỘNG TÀI SẢN' in latest_bs and pd.notna(
                        latest_bs['TỔNG CỘNG TÀI SẢN']) else None, "liabilities": float(latest_bs['NỢ PHẢI TRẢ']) if 'NỢ PHẢI TRẢ' in latest_bs and pd.notna(
                        latest_bs['NỢ PHẢI TRẢ']) else None, "current_liabilities": float(latest_bs['Nợ ngắn hạn']) if 'Nợ ngắn hạn' in latest_bs and pd.notna(
                        latest_bs['Nợ ngắn hạn']) else None, "non_current_liabilities": float(latest_bs['Nợ dài hạn']) if 'Nợ dài hạn' in latest_bs and pd.notna(
                        latest_bs['Nợ dài hạn']) else None,"equity": float(latest_bs['VỐN CHỦ SỞ HỮU']) if 'VỐN CHỦ SỞ HỮU' in latest_bs and pd.notna(
                        latest_bs['VỐN CHỦ SỞ HỮU']) else None, "owner_equity": float(
                        latest_bs['Vốn góp của chủ sở hữu']) if 'Vốn góp của chủ sở hữu' in latest_bs and pd.notna(
                        latest_bs['Vốn góp của chủ sở hữu']) else None, "retained_earnings": float(
                        latest_bs['Lãi chưa phân phối']) if 'Lãi chưa phân phối' in latest_bs and pd.notna(
                        latest_bs['Lãi chưa phân phối']) else None, "total_liabilities_and_equity": float(
                        latest_bs['TỔNG CỘNG NGUỒN VỐN']) if 'TỔNG CỘNG NGUỒN VỐN' in latest_bs and pd.notna(
                        latest_bs['TỔNG CỘNG NGUỒN VỐN']) else None
                }
        
        # Get income statement data
        if 'income_statement' in data:
            income_statement_data = data['income_statement'][
                (data['income_statement']['Mã'] == company_code) &
                (data['income_statement']['Năm'] == year)
                ]
            
            if not income_statement_data.empty:
                # Sort by quarter to get the latest quarter data
                income_statement_data = income_statement_data.sort_values(by='Quý', ascending=False)
                latest_is = income_statement_data.iloc[0]
                
                # Extract relevant fields
                result["income_statement"] = {
                    "total_revenue": float(latest_is['Doanh thu bán hàng và cung cấp dịch vụ']) if 'Doanh thu bán hàng và cung cấp dịch vụ' in latest_is and pd.notna(
                        latest_is['Doanh thu bán hàng và cung cấp dịch vụ']) else None, "revenue_deductions": float(latest_is['Doanh thu bán hàng và cung cấp dịch vụ']) - float(latest_is['Doanh thu thuần']) if 'Doanh thu bán hàng và cung cấp dịch vụ' in latest_is and 'Doanh thu thuần' in latest_is and pd.notna(
                        latest_is['Doanh thu bán hàng và cung cấp dịch vụ']) and pd.notna(
                        latest_is['Doanh thu thuần']) else None, "net_revenue": float(latest_is['Doanh thu thuần']) if 'Doanh thu thuần' in latest_is and pd.notna(
                        latest_is['Doanh thu thuần']) else None, "cost_of_goods_sold": float(latest_is['Doanh thu thuần']) - float(latest_is['Lợi nhuận gộp về bán hàng và cung cấp dịch vụ']) if 'Doanh thu thuần' in latest_is and 'Lợi nhuận gộp về bán hàng và cung cấp dịch vụ' in latest_is and pd.notna(
                        latest_is['Doanh thu thuần']) and pd.notna(
                        latest_is['Lợi nhuận gộp về bán hàng và cung cấp dịch vụ']) else None, "gross_profit": float(latest_is['Lợi nhuận gộp về bán hàng và cung cấp dịch vụ']) if 'Lợi nhuận gộp về bán hàng và cung cấp dịch vụ' in latest_is and pd.notna(
                        latest_is['Lợi nhuận gộp về bán hàng và cung cấp dịch vụ']) else None, "financial_income": float(latest_is['Doanh thu hoạt động tài chính']) if 'Doanh thu hoạt động tài chính' in latest_is and pd.notna(
                        latest_is['Doanh thu hoạt động tài chính']) else None, "financial_expenses": float(
                        latest_is['Chi phí tài chính']) if 'Chi phí tài chính' in latest_is and pd.notna(
                        latest_is['Chi phí tài chính']) else None, "interest_expense": float(latest_is['Trong đó: Chi phí lãi vay']) if 'Trong đó: Chi phí lãi vay' in latest_is and pd.notna(
                        latest_is['Trong đó: Chi phí lãi vay']) else None, "selling_expenses": float(
                        latest_is['Chi phí bán hàng']) if 'Chi phí bán hàng' in latest_is and pd.notna(
                        latest_is['Chi phí bán hàng']) else None, "administrative_expenses": float(latest_is['Chi phí quản lý doanh nghiệp']) if 'Chi phí quản lý doanh nghiệp' in latest_is and pd.notna(
                        latest_is['Chi phí quản lý doanh nghiệp']) else None, "operating_profit": float(latest_is['Lợi nhuận thuần từ hoạt động kinh doanh']) if 'Lợi nhuận thuần từ hoạt động kinh doanh' in latest_is and pd.notna(
                        latest_is['Lợi nhuận thuần từ hoạt động kinh doanh']) else None, 
                    "other_income": None, # Not directly available in the dataset
                    "other_expenses": None, # Not directly available in the dataset
                    "other_profit": float(latest_is['Lợi nhuận khác']) if 'Lợi nhuận khác' in latest_is and pd.notna(
                        latest_is['Lợi nhuận khác']) else None, "profit_before_tax": float(latest_is['Tổng lợi nhuận kế toán trước thuế']) if 'Tổng lợi nhuận kế toán trước thuế' in latest_is and pd.notna(
                        latest_is['Tổng lợi nhuận kế toán trước thuế']) else None, "current_tax": float(latest_is['Chi phí thuế thu nhập doanh nghiệp']) if 'Chi phí thuế thu nhập doanh nghiệp' in latest_is and pd.notna(
                        latest_is['Chi phí thuế thu nhập doanh nghiệp']) else None,
                    "deferred_tax": None, # Not directly available in the dataset
                    "profit_after_tax": float(latest_is['Lợi nhuận sau thuế thu nhập doanh nghiệp']) if 'Lợi nhuận sau thuế thu nhập doanh nghiệp' in latest_is and pd.notna(
                        latest_is['Lợi nhuận sau thuế thu nhập doanh nghiệp']) else None, "basic_earnings_per_share": float(
                        latest_is['Lãi cơ bản trên cổ phiếu']) if 'Lãi cơ bản trên cổ phiếu' in latest_is and pd.notna(
                        latest_is['Lãi cơ bản trên cổ phiếu']) else None
                }
        
        # Get cash flow statement data
        if 'cash_flow' in data:
            cash_flow_data = data['cash_flow'][
                (data['cash_flow']['Mã'] == company_code) &
                (data['cash_flow']['Năm'] == year)
                ]
            
            if not cash_flow_data.empty:
                # Sort by quarter to get the latest quarter data
                cash_flow_data = cash_flow_data.sort_values(by='Quý', ascending=False)
                latest_cf = cash_flow_data.iloc[0]
                
                # Extract relevant fields
                result["cash_flow"] = {
                    "profit_before_tax": float(latest_cf['Tổng lợi nhuận kế toán trước thuế.1']) if 'Tổng lợi nhuận kế toán trước thuế.1' in latest_cf and pd.notna(
                        latest_cf['Tổng lợi nhuận kế toán trước thuế.1']) else None, 
                    "adjustments": None, # Need to calculate from multiple fields
                    "depreciation": float(latest_cf['Khấu hao TSCĐ']) if 'Khấu hao TSCĐ' in latest_cf and pd.notna(
                        latest_cf['Khấu hao TSCĐ']) else None, "interest_expense": None, # Would need to crossreference with income statement
                        "net_cash_from_operating": float(latest_cf['Lưu chuyển tiền tệ ròng từ các hoạt động sản xuất kinh doanh (TT)']) if 'Lưu chuyển tiềntệ ròng từ các hoạt động sản xuất kinh doanh (TT)' in latest_cf and pd.notna(
                        latest_cf['Lưu chuyển tiền tệ ròng từ các hoạt động sản xuất kinh doanh (TT)']) else None, "purchase_of_fixed_assets": float(latest_cf['Tiền chi để mua sắm, xây dựng TSCĐ và các tài sản dài hạn khác (TT)']) if 'Tiền chi để mua sắm, xây dựng TSCĐ và các tài sản dài hạn khác (TT)' in latest_cf and pd.notna(
                        latest_cf['Tiền chi để mua sắm, xây dựng TSCĐ và các tài sản dài hạn khác (TT)']) else None, "proceeds_from_disposals": float(latest_cf['Tiền thu từ thanh lý, nhượng bán TSCĐ và các tài sản dài hạn khác (TT)']) if 'Tiền thu từ thanh lý, nhượng bán TSCĐ và các tài sản dài hạn khác (TT)' in latest_cf and pd.notna(
                        latest_cf['Tiền thu từ thanh lý, nhượng bán TSCĐ và các tài sản dài hạn khác (TT)']) else None, "loans_to_other_entities": float(latest_cf['Tiền chi cho vay, mua các công cụ nợ của đợn vị khác (TT)']) if 'Tiền chi cho vay, mua các công cụ nợ của đợn vị khác (TT)' in latest_cf and pd.notna(
                        latest_cf['Tiền chi cho vay, mua các công cụ nợ của đợn vị khác (TT)']) else None, "collections_from_loans": float(latest_cf['Tiền thu hồi cho vay, bán lại các công cụ nợ của đơn vị khác (TT)']) if 'Tiền thu hồi cho vay, bán lại các công cụ nợ của đơn vị khác (TT)' in latest_cf and pd.notna(
                        latest_cf['Tiền thu hồi cho vay, bán lại các công cụ nợ của đơn vị khác (TT)']) else None, "net_cash_from_investing": float(latest_cf['Lưu chuyển tiền tệ ròng từ hoạt động đầu tư (TT)']) if 'Lưu chuyển tiền tệ ròng từ hoạt động đầu tư (TT)' in latest_cf and pd.notna(
                        latest_cf['Lưu chuyển tiền tệ ròng từ hoạt động đầu tư (TT)']) else None, "proceeds_from_issuing_shares": float(latest_cf['Tiền thu từ phát hành cổ phiếu, nhận góp vốn của chủ sở hữu (TT)']) if 'Tiền thu từ phát hành cổ phiếu, nhận góp vốn của chủ sở hữu (TT)' in latest_cf and pd.notna(
                        latest_cf['Tiền thu từ phát hành cổ phiếu, nhận góp vốn của chủ sở hữu (TT)']) else None, "proceeds_from_borrowings": float(latest_cf['Tiền thu được các khoản đi vay (TT)']) if 'Tiền thu được các khoản đi vay (TT)' in latest_cf and pd.notna(
                        latest_cf['Tiền thu được các khoản đi vay (TT)']) else None, "repayments_of_borrowings": float(
                        latest_cf['Tiền trả nợ gốc vay (TT)']) if 'Tiền trả nợ gốc vay (TT)' in latest_cf and pd.notna(
                        latest_cf['Tiền trả nợ gốc vay (TT)']) else None,"dividends_paid": float(
                        latest_cf['Cổ tức đã trả (TT)']) if 'Cổ tức đã trả (TT)' in latest_cf and pd.notna(
                        latest_cf['Cổ tức đã trả (TT)']) else None, "net_cash_from_financing": float(latest_cf['Lưu chuyển tiền tệ từ hoạt động tài chính (TT)']) if 'Lưu chuyển tiền tệ từ hoạt động tài chính (TT)' in latest_cf and pd.notna(
                        latest_cf['Lưu chuyển tiền tệ từ hoạt động tài chính (TT)']) else None, "net_cash_flow": float(latest_cf['Lưu chuyển tiền thuần trong kỳ (TT)']) if 'Lưu chuyển tiền thuần trong kỳ (TT)' in latest_cf and pd.notna(
                        latest_cf['Lưu chuyển tiền thuần trong kỳ (TT)']) else None, "cash_beginning": float(latest_cf['Tiền và tương đương tiền đầu kỳ (TT)']) if 'Tiền và tương đương tiền đầu kỳ (TT)' in latest_cf and pd.notna(
                        latest_cf['Tiền và tương đương tiền đầu kỳ (TT)']) else None, "cash_ending": float(latest_cf['Tiền và tương đương tiền đầu kỳ (TT)']) + float(latest_cf['Lưu chuyển tiền thuần trong kỳ (TT)']) if 'Tiền và tương đương tiền đầu kỳ (TT)' in latest_cf and 'Lưu chuyển tiền thuần trong kỳ (TT)' in latest_cf and pd.notna(
                        latest_cf['Tiền và tương đương tiền đầu kỳ (TT)']) and pd.notna(
                        latest_cf['Lưu chuyển tiền thuần trong kỳ (TT)']) else None
                }
        
        return result
    
    except Exception as e:
        print(f"Error fetching financial statements for {company_code}, year {year}: {e}")
        return {"error": str(e)}

# Hàm xuất dữ liệu ra PDF
def export_to_pdf(data, charts=None, filename='financial_report.pdf'):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    
    # Add title
    pdf.set_font("Arial", style="B", size=16)
    pdf.cell(200, 10, txt="Báo cáo tài chính", ln=True, align="C")
    pdf.ln(10)
    
    # Add content from datapdf.set_font("Arial", size=12)
    for key, value in data.items():
        pdf.cell(0, 10, txt=f"{key}: {str(value)}", ln=True)
    
    # Add charts to PDF
    if charts:
        valid_charts = [chart for chart in charts if chart is not None]
# Filter out None values
    for chart in valid_charts:
        try:
            pdf.add_page()
            pdf.image(chart, x=10, y=20, w=180)
        except Exception as e:
            print(f"Error adding chart to PDF: {e}")

# Save PDF
pdf.output(filename)
print(f"✅ Đã xuất báo cáo ra file: {filename}")

    # === Chương trình chính ===
if __name__ == "__main__":
    # Đường dẫn thư mục dữ liệu
    base_path = 'C:/Users/Dell/OneDrive/Tài liệu/GÓI 1/127 hoan chinh/data'
    output_dir = 'C:/Users/Dell/OneDrive/Tài liệu/GÓI 1/127 hoan chinh/output'
    
    # Tạo thư mục output nếu chưa tồn tại
    os.makedirs(output_dir, exist_ok=True)
    
    # Load dữ liệu
    data = load_data()
    
    # Xử lý dữ liệu cho công ty cụ thể
    company_code = "MWG" # Thay bằng mã công ty bạn muốn
    company_data = get_company_report_data(company_code)
    
    if company_data:
        # Tạo biểu đồ
        prepare_financial_charts(company_data)
        charts = [
            company_data.get('revenue_profit_chart'),
            company_data.get('ratios_chart'),
            company_data.get('balance_sheet_chart'),
            company_data.get('comparison_chart'),
            company_data.get('profitability_chart'),
            company_data.get('growth_chart'),
            company_data.get('liquidity_chart'),
            company_data.get('leverage_chart'),
            company_data.get('efficiency_chart')
        ]
        
        # Xuất báo cáo PDF
        pdf_filename = os.path.join(output_dir,f"{company_code}_financial_report.pdf")
        export_to_pdf(company_data, charts, pdf_filename)