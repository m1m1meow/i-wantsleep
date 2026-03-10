import pandas as pd

def convert_txt_to_xlsx(input_file, output_file):
    data = []
    try:
        with open(input_file, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                
                # maxsplit=2 代表切兩刀，分成三部分：時間、內容、講者
                parts = line.split(maxsplit=2)
                
                # 確保資料長度一致，避免程式崩潰
                if len(parts) == 3:
                    data.append(parts)
                elif len(parts) == 2:
                    # 如果只有時間和內容，講者留空
                    data.append([parts[0], parts[1], ""])
                else:
                    # 如果只有時間，其餘留空
                    data.append([parts[0], "", ""])

        # 建立 DataFrame，加入「講者」欄位
        df = pd.DataFrame(data, columns=['時間', '內容', '講者'])
        
        # 儲存為 Excel
        df.to_excel(output_file, index=False, engine='openpyxl')
        
        print(f"轉換成功！已新增講者欄位，檔案：{output_file}")

    except Exception as e:
        print(f"發生錯誤：{e}")

# 執行轉換
convert_txt_to_xlsx('workshop_agenda.txt', 'workshop_final.xlsx')
