def k5(f2, f3):
    import pymysql
    import xlwt
    from sqlalchemy import create_engine
    import xlwings as xw
    import pandas as pd

    # 第二轮数据处理
    conn = pymysql.connect(
        host='localhost',
        port=3306,
        user='root',
        password='root',
        db="excel",
        charset='utf8'
    )
    engine = create_engine('mysql+pymysql://root:root@localhost/excel')
    df2 = pd.read_excel(f2)
    df2.to_sql(name='second', con=engine, index=False, if_exists='replace')
    df3 = pd.read_excel(f3)
    df3.to_sql(name='required', con=engine, index=False, if_exists='replace')
    cur = conn.cursor()
    sql = 'SELECT s.学号,s.姓名,s.不及格课程名称,s.成绩,s.学分 FROM second AS s,required AS r WHERE s.不及格课程名称 = r.课程名称 ORDER BY 学号;'
    wb_new = xlwt.Workbook(f2)
    sht_new = wb_new.add_sheet('data')
    sht_new.write(0, 0, '学号')
    sht_new.write(0, 1, '姓名')
    sht_new.write(0, 2, '必修课名称')
    sht_new.write(0, 3, '成绩')
    sht_new.write(0, 4, '学分')

    try:
        cur.execute(sql)
        sales = cur.fetchall()
        n = 1
        print(sales)
        for sale in sales:
            a = sale[0]
            b = sale[1]
            c = sale[2]
            d = sale[3]
            e = sale[4]

            sht_new.write(n, 0, a)
            sht_new.write(n, 1, b)
            sht_new.write(n, 2, c)
            sht_new.write(n, 3, d)
            sht_new.write(n, 4, e)

            n += 1
            wb_new.save(f2)
            print('获取数据成功！')
        app = xw.App(visible=True, add_book=False)
        app.books.open(f2)
    except Exception as e:
        print(e)
        print('导出数据失败！')
    finally:
        conn.close()
        cur.close()
