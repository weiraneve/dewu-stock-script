import openpyxl

def read_stock(file_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    # 定位数据起始行，假设销售订单页的实际数据从第2行开始
    data_start_row = 2

    data = []
    for row in sheet.iter_rows(min_row=data_start_row, values_only=True):
        if row[0] is None:  # 如果第一列为空，跳过无效行
            continue
        data.append({
            "仓库": row[0],
            "商品名称": row[1],
            "货号": row[2],
            "尺码": row[3],
            "成本价": row[4],
            "库存": row[5],
            "当前毒普通价": row[6],
            "价格更新时间": row[7],
            "3.5到手": row[8],
            "4.0到手": row[9],
            "5.0到手": row[10],
            "入库时间": row[11],
            "备注": row[12],
        })
    return data

def read_dewu(file_path):
    # 打开工作簿
    workbook = openpyxl.load_workbook(file_path)
    
    # 固定指定的工作表为“销售订单”
    sheet_name = "销售订单"
    if sheet_name not in workbook.sheetnames:
        raise ValueError(f"工作表 '{sheet_name}' 不存在！有效的工作表为: {workbook.sheetnames}")
    
    # 加载“销售订单”工作表
    sheet = workbook[sheet_name]

    # 定位数据起始行，假设实际数据从第4行开始
    data_start_row = 4

    data = []
    for row in sheet.iter_rows(min_row=data_start_row, values_only=True):
        # 确保行中列数足够并且关键列不为空
        if len(row) < 58 or row[3] is None or row[5] is None or row[57] is None:
            continue
        data.append({
            "商品货号": row[3],  # D列
            "规格": row[5],       # F列
            "实付金额": row[57],   # BG列
        })
    return data

def compare_and_calculate(data_stock, data_dewu):
    results = []
    for row_stock in data_stock:
        for row_dewu in data_dewu:
            if row_stock["货号"] == row_dewu["商品货号"] and row_stock["尺码"] == row_dewu["规格"]:
                cost_price = row_stock["成本价"]
                paid_amount = row_dewu["实付金额"]
                print(paid_amount)
                difference = round(float(paid_amount)) - round(float(cost_price))
                results.append({
                    "货号": row_stock["货号"],
                    "尺码": row_stock["尺码"],
                    "成本价": cost_price,
                    "实付金额": paid_amount,
                    "利润": difference,
                })
    return results

def main():
    stock_path = "瑕疵成本对照7.11.xlsx"
    dewu_path = "24.11.22-12.20大雄得物.xlsx"

    data_stock = read_stock(stock_path)
    data_dewu = read_dewu(dewu_path)

    results = compare_and_calculate(data_stock, data_dewu)

    for result in results:
        print(result)

if __name__ == "__main__":
    main()