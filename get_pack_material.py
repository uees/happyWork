import argparse

from package_material.parser import WorksheetParser

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='分析包材用量')
    parser.add_argument("excel_file", help="文件名")
    parser.add_argument("-i", "--index", type=int, default=2, help="excel中需要生成报告的起始行")
    parser.add_argument("-s", "--sheet", default="产品进仓", help="excel中的工作表")

    args = parser.parse_args()
    parser = WorksheetParser(args.excel_file, args.sheet, args.index)
    parser.run()
