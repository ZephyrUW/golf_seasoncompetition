import os
import sqlite3
import shutil
from datetime import datetime, date

class Database:
    def __init__(self, db):
        try:
            self.conn = sqlite3.connect(db, detect_types=sqlite3.PARSE_DECLTYPES)
        except sqlite3.OperationalError:
            try:
                self.conn = sqlite3.connect(f"./{db}", detect_types=sqlite3.PARSE_DECLTYPES)
            except sqlite3.OperationalError:
                print(f"{os.path.dirname(__file__)}{db}")
                self.conn = sqlite3.connect(f"{os.path.dirname(__file__)}\\{db}", detect_types=sqlite3.PARSE_DECLTYPES)



        self.cur = self.conn.cursor()
        self.cur.execute(
            "CREATE TABLE IF NOT EXISTS housecharges (id INTEGER PRIMARY KEY, chargedate date, employee text, amt int)")
        self.conn.commit()
        if not (date.today().day % 3):
            self.backup(db)

    def fetch_group(self, startdate, enddate=date.today()):
        if type(startdate) is str:
            startdate = date.fromisoformat(startdate)
        if type(enddate) is str:
            enddate = date.fromisoformat(enddate)

        # todo Need to add chargedate ranges
        self.cur.execute("SELECT * FROM housecharges WHERE chargedate BETWEEN ? AND ?", (startdate, enddate, ))
        rows = self.cur.fetchall()
        # todo return sums?
        return rows

    def fetch_one(self, name, startdate, enddate):
        self.cur.execute("SELECT * FROM housecharges WHERE employee=? AND chargedate BETWEEN ? AND ?", (name, startdate, enddate))
        rows = self.cur.fetchall()
        return rows

    def insert(self, chargedate, hc):
        print(hc)
        temp_hc = [[chargedate, i[0], int(i[1]*100)] for i in hc]
        self.cur.executemany("INSERT INTO housecharges VALUES (NULL, ?, ?, ?)", temp_hc)
        self.conn.commit()
        
        return len(temp_hc)

    def remove(self, id):
        self.cur.execute("DELETE FROM housecharges WHERE id=?", (id,))
        self.conn.commit()

    def update(self, id, chargedate, employee, amt):
        self.cur.execute("UPDATE housecharges SET chargedate = ?, employee = ?, amt = ?, WHERE id = ?",
                         (date, employee, amt*100, id))
        self.conn.commit()

    def backup(self, db):
        shutil.copyfile(db, f"{os.path.dirname(__file__)}\\data\\backup\\house_charges_{str(date.today())}.db")

    def all_employee_names(self):
        self.cur.execute("SELECT DISTINCT employee FROM housecharges ORDER BY employee")
        employee_list = self.cur.fetchall()

        return employee_list


    def manual_entry(self, filename):
        wb = openpyxl.load_workbook(filename, keep_vba=True, read_only=True, data_only=True)
        s = wb['Rankings']

        # Create list of name and amount cells
        # name_cells = [f'I{i}' for i in range(5,10)]
        # name_cells.extend([f'K{i}' for i in range(4,10)])
        # name_cells.extend([f'M{i}' for i in range(4,10)])
        # amt_cells = [f'J{i}' for i in range(5,10)]
        # amt_cells.extend([f'L{i}' for i in range(4,10)])
        # amt_cells.extend([f'N{i}' for i in range(4,10)])
        
        #     # Get House Charges
        # for i in range(len(name_cells)):
        #     if s[amt_cells[i]].value == '' or s[amt_cells[i]].value == None or s[amt_cells[i]].value == ' ':
        #         continue
        #     hc.append([str(s[name_cells[i]].value).strip().title(), s[amt_cells[i]].value])
            
        wb.close()
        # print((filename.split('.')[0].split('\\')[-1]))
        
        # _ = self.insert(date.fromisoformat(filename.split('.')[0].split('\\')[-1]), hc)

    def __del__(self):
        try:
            self.conn.close()
        except:
            pass

if __name__ == '__main__':
    import os
    import openpyxl
    
    # TODO Change to Server
    # '\\\\server\MorningBooks\data'
    db = Database('golf_results.db')
    folder = '\\\\DESKTOP-06TLJUH\\Users\\PIN AND CUE\\Desktop\\Shared Drive\\Scott\\Golf\\'
    
    ## TESTING ##
    folder = "./test_results"
    # Fix for no network access...unkown reason why...
    # folder = "Z:/"

    files = os.listdir(folder)
    
    golf_files = [i for i in files if ".xlsm" in i]
    print(files)
    print(golf_files)
    
    start = int(input('Type ENTER to import all files listed above\n'))
    
    for excel in golf_files:
        db.manual_entry(f'{folder}{excel}')
            