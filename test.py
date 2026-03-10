import csv

def convert_txt_to_csv(input_file, output_file):
    try:
        # 使用 utf-8 讀取，避免繁體中文亂碼
        with open(input_file, 'r', encoding='utf-8') as txtfile:
            lines = txtfile.readlines()

        with open(output_file, 'w', newline='', encoding='utf-8-sig') as csvfile:
            writer = csv.writer(csvfile)
            # 寫入標題列
            writer.writerow(['時間', '內容'])

            for line in lines:
                # 移除行尾換行符號與前後空白
                line = line.strip()
                if not line:
                    continue  # 跳過空行

                # 假設時間與內容之間是用空格或 Tab 分隔
                # 如果你的格式更複雜（例如：09:00~10:00 這是內容）
                # 我們可以用 split(maxsplit=1) 拆分前兩個部分
                parts = line.split(maxsplit=1)
                
                if len(parts) == 2:
                    writer.writerow([parts[0], parts[1]])
                else:
                    # 如果只有時間沒有內容，則補空值
                    writer.writerow([parts[0], ""])

        print(f"成功！已轉換為: {output_file}")

    except FileNotFoundError:
        print("錯誤：找不到來源檔案。")
    except Exception as e:
        print(f"發生錯誤：{e}")

# 執行轉換
convert_txt_to_csv('input.txt', 'output.csv')
