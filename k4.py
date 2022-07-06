def k4(f1, f2):
    import pandas as pd
    import pymysql
    import xlwt
    from sqlalchemy import create_engine
    import xlwings as xw

    conn = pymysql.connect(
            host='localhost',
            port=3306,
            user='root',
            password='root',
            db="excel",
            charset='utf8'
    )
    engine = create_engine('mysql+pymysql://root:root@localhost/excel')
    df1 = pd.read_excel(f1)
    df1.to_sql(name='original', con=engine, index=False, if_exists='replace', chunksize=8888)
    cur = conn.cursor()
    sql = 'SELECT 学号,姓名,课程名称,学分 FROM original WHERE 课程性质 LIKE "%公共选修%"' \
          'ORDER BY 学号;'

    wb_new = xlwt.Workbook(f2)
    sht_new = wb_new.add_sheet('data')
    sht_new.write(0, 0, '学号')
    sht_new.write(0, 1, '姓名')
    sht_new.write(0, 2, '公选课名称')
    sht_new.write(0, 3, '学分')

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
            sht_new.write(n, 0, a)
            sht_new.write(n, 1, b)
            sht_new.write(n, 2, c)
            sht_new.write(n, 3, d)
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
