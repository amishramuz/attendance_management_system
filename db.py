import sqlite3


class Database:
    def __init__(self,db):
        self.conn=sqlite3.connect(db)
        self.cur=self.conn.cursor()
        self.cur.execute("CREATE TABLE IF NOT EXISTS STUDENT (id integer primary key,name text,regi INTEGER ,branch text)")
        self.conn.commit()

    def fetch(self):
        self.cur.execute("select * from STUDENT")
        rows=self.cur.fetchall()
        return rows

    def insert(self,name,regi,branch):
        self.cur.execute("INSERT into STUDENT values (null,?,?,?)",(name,regi,branch))
        self.conn.commit()

    def remove(self,regi):
        self.cur.execute("delete from STUDENT where regi=(?)",(regi))
        self.conn.commit()
    
    
    

    def __del__(self):
        self.conn.close()


