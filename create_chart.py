
import os
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime

def get_today_production_data():
    today_str = datetime.now().strftime('%Y-%m-%d')
    all_data = []
    
    for root, _, files in os.walk('csv_output'):
        for file in files:
            if file.endswith('.csv'):
                file_path = os.path.join(root, file)
                try:
                    df = pd.read_csv(file_path, low_memory=False)
                    # This is a guess, you might need to adjust column names
                    date_col = next((col for col in df.columns if 'date' in col.lower() or '時間' in col.lower() or 'Unnamed: 12' in col), None)
                    name_col = next((col for col in df.columns if '品名' in col or 'product' in col.lower() or 'Unnamed: 4' in col), None)
                    qty_col = next((col for col in df.columns if '數量' in col or 'quantity' in col.lower() or 'Unnamed: 14' in col), None)

                    if date_col and name_col and qty_col:
                        # Coerce to datetime, handling errors by turning problematic entries into NaT
                        df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
                        # Drop rows where date could not be parsed
                        df.dropna(subset=[date_col], inplace=True)
                        
                        today_df = df[df[date_col].dt.strftime('%Y-%m-%d') == today_str].copy()

                        if not today_df.empty:
                            today_df[qty_col] = pd.to_numeric(today_df[qty_col], errors='coerce').fillna(0)
                            all_data.append(today_df[[name_col, qty_col]])
                except Exception as e:
                    print(f"Error processing {file_path}: {e}")

    if not all_data:
        return pd.DataFrame()

    combined_df = pd.concat(all_data)
    name_col_agg = combined_df.columns[0]
    qty_col_agg = combined_df.columns[1]
    
    product_counts = combined_df.groupby(name_col_agg)[qty_col_agg].sum().sort_values(ascending=False)
    return product_counts

def create_production_chart(product_counts):
    if product_counts.empty:
        print("No production data found for today.")
        return

    top_15 = product_counts.head(15)

    plt.figure(figsize=(12, 8))
    top_15.plot(kind='bar')
    plt.title(f"Top 15 Products by Production Quantity - {datetime.now().strftime('%Y-%m-%d')}")
    plt.xlabel("Product Name")
    plt.ylabel("Total Quantity")
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    
    chart_path = 'production_chart.png'
    plt.savefig(chart_path)
    print(f"Chart saved to {os.path.abspath(chart_path)}")

if __name__ == "__main__":
    # To handle potential font issues with Chinese characters
    plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei'] 
    plt.rcParams['axes.unicode_minus'] = False

    production_data = get_today_production_data()
    create_production_chart(production_data)
