import re
import openpyxl
from openpyxl import Workbook

def read_stock(file_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    data_start_row = 2
    data = []
    for row in sheet.iter_rows(min_row=data_start_row, values_only=True):
        if row[0] is None:
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
    workbook = openpyxl.load_workbook(file_path)
    sheet_name = "销售订单"
    if sheet_name not in workbook.sheetnames:
        raise ValueError(f"工作表 '{sheet_name}' 不存在！有效的工作表为: {workbook.sheetnames}")
    sheet = workbook[sheet_name]

    data_start_row = 4
    data = []
    for row in sheet.iter_rows(min_row=data_start_row, values_only=True):
        if len(row) < 58 or row[3] is None or row[5] is None or row[57] is None:
            continue
        data.append({
            "商品货号": row[3],
            "数量": row[4],
            "规格": row[5],
            "实付金额": row[57],
        })
    return data

def compare_and_calculate(data_stock, data_dewu):
    # 创建一个字典来存储得物订单，键为(货号, 规格)组合
    dewu_dict = {}
    for row_dewu in data_dewu:
        # 提取规格中的数字并转换为字符串
        size = re.search(r'\d+\.?\d*', str(row_dewu["规格"]))
        size = size.group() if size else str(row_dewu["规格"])
        key = (row_dewu["商品货号"], size)
        if key not in dewu_dict:
            dewu_dict[key] = []
        dewu_dict[key].append(row_dewu)

    results = []
    for row_stock in data_stock:
        # 确保库存中的尺码也是字符串
        stock_size = str(row_stock["尺码"])
        key = (row_stock["货号"], stock_size)
        
        if key in dewu_dict:
            stock = int(row_stock["库存"]) if row_stock["库存"] is not None else 0
            remaining_stock = stock

            # 处理该商品的所有订单
            for row_dewu in dewu_dict[key]:
                cost_price = float(row_stock["成本价"]) if row_stock["成本价"] is not None else 0
                paid_amount = float(row_dewu["实付金额"]) if row_dewu["实付金额"] is not None else 0
                
                difference = round(paid_amount) - round(cost_price)
                sold_quantity = min(row_dewu["数量"], remaining_stock)
                remaining_stock -= sold_quantity

                if sold_quantity > 0:
                    result = row_stock.copy()
                    result["利润"] = difference
                    result["库存"] = remaining_stock
                    result["卖出数量"] = sold_quantity
                    results.append(result)

                if remaining_stock <= 0:
                    break

    return results

def write_to_excel(data, headers, output_path, include_profit):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "计算结果" if include_profit else "导入文件"

    # 写入表头
    if include_profit:
        sheet.append(headers + ["利润", "卖出数量"])
    else:
        sheet.append(headers)

    # 写入数据
    for row_data in data:
        row = [row_data.get(header, "") for header in headers]
        if include_profit:
            row.append(row_data.get("利润", ""))
            row.append(row_data.get("卖出数量", ""))
        sheet.append(row)

    workbook.save(output_path)

def main():
    stock_path = "瑕疵成本对照7.11.xlsx"
    dewu_path = "24.11.22-12.20大雄得物.xlsx"
    output_profit_path = "利润结果.xlsx"
    output_import_path = "导入文件.xlsx"

    data_stock = read_stock(stock_path)
    data_dewu = read_dewu(dewu_path)

    headers = ["仓库", "商品名称", "货号", "尺码", "成本价", "库存", "当前毒普通价", "价格更新时间", "3.5到手", "4.0到手", "5.0到手", "入库时间", "备注"]

    results = compare_and_calculate(data_stock, data_dewu)

    # 生成“利润结果”文件（包含利润和库存列）
    write_to_excel(results, headers, output_profit_path, include_profit=True)

    # 生成“导入文件”文件（不包含利润列，但库存已更新）
    write_to_excel(results, headers, output_import_path, include_profit=False)

    print(f"文件已生成：\n- {output_profit_path}（包含利润和库存）\n- {output_import_path}（库存已更新，无利润列）")

if __name__ == "__main__":
    main()
