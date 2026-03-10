import pandas as pd
from openpyxl.utils import get_column_letter

def convert_txt_to_xlsx(input_file, output_file):
    data = []
    try:
        with open(input_file, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue

                # 這裡假設你的格式是: "時間 內容 講者" 或是 "時間,內容,講者"
                # 如果你是用 tab 隔開，請將 split() 改成 split('\t')
                parts = line.split(maxsplit=2)

                # 確保長度補齊到 3 欄
                while len(parts) < 3:
                    parts.append("")
                
                data.append(parts[:3]) # 只取前三欄，避免超出

        # 1. 建立 DataFrame
        df = pd.DataFrame(data, columns=['時間', '內容', '講者'])

        # 2. 建立 Excel 寫入器，使用 xlsxwriter 引擎可以做更多格式設定
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='工作坊議程')
            
            # 3. 自動調整欄位寬度 (美化輸出)
            worksheet = writer.sheets['工作坊議程']
            for i, col in enumerate(df.columns):
                # 計算該欄位中最長字串的長度
                max_len = max(
                    df[col].astype(str).map(len).max(),
                    len(col)
                ) + 5  # 額外加 5 個字元的餘裕
                worksheet.column_dimensions[get_column_letter(i+1)].width = max_len

        print(f" 轉換成功！檔案已美化：{output_file}")

    except Exception as e:
        print(f" 發生錯誤：{e}")

# 執行轉換
convert_txt_to_xlsx('workshop_agenda.txt', 'workshop_final.xlsx')
