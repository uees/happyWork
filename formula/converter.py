import platform

if platform.system() == 'Windows':

    from win32com.client import constants, Dispatch
    from pywintypes import com_error

    def convert2xlsx(files):
        excel = Dispatch("Excel.Application")
        excel.Visible = False

        for filename in files:
            print(f"转化 {filename}")
            try:
                workbook = excel.Workbooks.Open(filename)

                new_filename = filename.replace('工单', 'work_orders')
                dir_name = os.path.dirname(new_filename)
                if not os.path.exists(dir_name):
                    os.makedirs(dir_name)

                # xlWorkbookDefault = 51 表示用Excel2007或2010的格式（*.xlsx）来储存
                workbook.SaveAs(Filename=new_filename,
                                FileFormat=constants.xlWorkbookDefault)
                workbook.Close()
                print(f"转化成功 {new_filename}")
            except com_error:
                continue

        excel.Quit()

else:
    def convert2xlsx():
        pass


if __name__ == "__main__":
    import os
    import glob

    from settings import BASE_DIR

    WORK_ORDERS_PATH = os.path.join(BASE_DIR, os.path.join('data', '工单'))
    all_formula_files = glob.glob(f"{WORK_ORDERS_PATH}/**/*.xlsx", recursive=True)
    convert2xlsx(all_formula_files)
