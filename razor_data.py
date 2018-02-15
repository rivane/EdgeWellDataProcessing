import xlrd
import os
import pandas as pd
import zipfile


def process_tmall_data():
    raw_data_dir = r'C:\Users\Administrator\Downloads'
    files = os.listdir(raw_data_dir)

    # df_dict = {}
    rows = []
    for f in files:
        if f.endswith('xls'):
            rb = xlrd.open_workbook(os.path.join(raw_data_dir,f))
            rs = rb.sheet_by_index(0)
            n = rs.nrows

            dateCols = [f[11:21]]* (n-4)
            dateCols.insert(0,'date')
            platCols = ['Tmall']*(n-4)
            platCols.insert(0,'plaform')
            idCols = rs.col_values(1,start_rowx=3)
            titileCols = rs.col_values(2,start_rowx=3)
            linkCols = rs.col_values(4,start_rowx=3)
            uvCols = rs.col_values(6,start_rowx=3)
            pvCols = rs.col_values(5,start_rowx=3)
            salesCols = rs.col_values(15,start_rowx=3)
            paidCusCols = rs.col_values(27,start_rowx=3)
            qtySKUCols = rs.col_values(16,start_rowx=3)
            addFavorCols = rs.col_values(22,start_rowx=3)
            addToCartCols = rs.col_values(17,start_rowx=3)
            stayTimeCols = rs.col_values(7,start_rowx=3)
            bounceRateCols = rs.col_values(8,start_rowx=3)
            cvrCols = rs.col_values(11,start_rowx=3)
            atvCols = rs.col_values(24,start_rowx=3)
            rvCols = rs.col_values(28,start_rowx=3)
            rtCols = rs.col_values(29,start_rowx=3)

            columns = [dateCols[0],platCols[0],idCols[0],titileCols[0],linkCols[0],uvCols[0],pvCols[0],salesCols[0],
                       paidCusCols[0],qtySKUCols[0],addFavorCols[0],addToCartCols[0],stayTimeCols[0],
                       bounceRateCols[0],cvrCols[0],atvCols[0],rvCols[0],rtCols[0]]

            for i in range(1,len(idCols)):
                row = [dateCols[i],platCols[i],idCols[i],titileCols[i],linkCols[i],uvCols[i],pvCols[i],salesCols[i],
                       paidCusCols[i],qtySKUCols[i],addFavorCols[i],addToCartCols[i],stayTimeCols[i],
                       bounceRateCols[i],cvrCols[i],atvCols[i],rvCols[i],rtCols[i]]
                rows.append(row)


    df = pd.DataFrame(rows,columns=columns)
    df.to_excel('razorData.xlsx',index=False,index_label=False)




def zipJDzipFile():
    raw_data_dir = r'C:\Users\Administrator\Downloads'
    files = os.listdir(raw_data_dir)

    for f in files:
        if f.endswith('zip'):
            file_zip = zipfile.ZipFile(os.path.join(raw_data_dir, f), 'r')
            for zf in file_zip.namelist():
                file_zip.extract(zf, raw_data_dir)
            file_zip.close()


def process_jd_razor_data():
    raw_data_dir = r'C:\Users\Administrator\Downloads'
    files = os.listdir(raw_data_dir)

    rows = []
    for f in files:
        if f.endswith('xls'):
            rb = xlrd.open_workbook(os.path.join(raw_data_dir, f))
            rs = rb.sheet_by_index(0)
            n = rs.nrows
            dates = f.split('_')[1]
            dates = '-'.join([dates[:4],dates[4:6],dates[6:]])

            dateCols = [dates] * (n - 2)
            platCols = ['JD'] * (n - 2)
            idCols = rs.col_values(0, start_rowx=2)
            titileCols = rs.col_values(1, start_rowx=2)
            linkCols = ['https://mall.jd.com/view_search-893764-0-5-1-24-1.html'] * (n - 2)
            uvCols = rs.col_values(3, start_rowx=2)
            pvCols = rs.col_values(4, start_rowx=2)
            salesCols = rs.col_values(11, start_rowx=2)
            orderCols = rs.col_values(9,start_rowx=2)

            paidCusCols = rs.col_values(8, start_rowx=2)
            qtySKUCols = rs.col_values(10, start_rowx=2)
            addFavorCols = rs.col_values(5, start_rowx=2)
            addToCartCols = rs.col_values(6, start_rowx=2)
            stayTimeCols = [0]* (n - 2)
            bounceRateCols = [0]* (n - 2)
            cvrCols = rs.col_values(12, start_rowx=2)
            atvCols = []
            for i in range(n - 2):
                if salesCols[i] > 0 :
                    atvCols.append(salesCols[i] / orderCols[i])
                else:
                    atvCols.append(0)

            rvCols = [0]* (n - 2)
            rtCols = [0]* (n - 2)

            for i in range(0,len(idCols)):
                row = [dateCols[i],platCols[i],idCols[i],titileCols[i],linkCols[i],uvCols[i],pvCols[i],salesCols[i],
                       paidCusCols[i],qtySKUCols[i],addFavorCols[i],addToCartCols[i],stayTimeCols[i],
                       bounceRateCols[i],cvrCols[i],atvCols[i],rvCols[i],rtCols[i]]
                rows.append(row)
    columns = ['date', 'plaform', '商品id', '商品标题', '商品链接', '访客数', '浏览量', '支付金额', '支付买家数', '支付商品件数',
               '收藏人数', '加购件数', '平均停留时长', '详情页跳出率', '支付转化率', '客单价', '售中售后成功退款金额', '售中售后成功退款笔数']

    df1 = pd.DataFrame(rows, columns=columns)
    if os.path.exists('razorData.xlsx'):
        df = pd.read_excel('razorData.xlsx')
        print(len(df.index))
        df = df.append(df1)
        print(len(df.index))
        df.to_excel('razorData.xlsx', index=False, index_label=False)



if __name__ == '__main__':
    process_jd_razor_data()


