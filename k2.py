def k2(f1, f2):
    import pandas as pd
    import pymysql
    import xlwt
    from sqlalchemy import create_engine

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

    # 第1轮数据处理
    sql = 'SELECT 学号,姓名,课程名称,成绩,学分 FROM original WHERE 成绩 < "60" and 成绩 != "100";'
    wb_new = xlwt.Workbook(f2)
    sht_new = wb_new.add_sheet('second')
    sht_new.write(0, 0, '学号')
    sht_new.write(0, 1, '姓名')
    sht_new.write(0, 2, '不及格课程名称')
    sht_new.write(0, 3, '成绩')
    sht_new.write(0, 4, '学分')
    try:
        cur.execute(sql)
        print(2)
        sales = cur.fetchall()
        print(1)
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
    except Exception as e:
        print(e)
        print('导出数据失败！')
    finally:
        conn.close()
        cur.close()
