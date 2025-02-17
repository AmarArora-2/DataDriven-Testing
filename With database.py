insert ="insert into student value(184,'Kim',89)"
update = "update student set name='Mary' where id=1"
delete = "delete from student where id=2"
import pymysql
try:
    con = pymysql.connect(
        host="localhost",
        port=3306,
        user="root",
        passwd="amar@123",
        database="youtube"
    )
    curs = con.cursor()
    print(curs)
    curs.execute(delete)
    con.commit()
    curs.execute("select * from student")
    for row in curs:
        print(row[0],row[1],row[2])

    con.close()
    print("finshed.....")
except:
    print("database not connected.....")

    # curs.execute("select * from students")
    # for row in curs:
    #     print(row[0],row[1],row[2])

