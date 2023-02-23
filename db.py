import os
import sqlite3
import shutil
from datetime import datetime, date

from rich import print, traceback
traceback.install()

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
            "CREATE TABLE IF NOT EXISTS golfresults (id INTEGER PRIMARY KEY, tournament text, entrant text, scoretotal int, placement int, junk int, importdate text)")
        self.cur.execute(
            "CREATE TABLE IF NOT EXISTS seasonresults (id INTEGER PRIMARY KEY, season text, entrant text, points int)")
        self.conn.commit()
        if not (date.today().day % 7):
            self.backup(db)

    def fetch_group(self, startdate, enddate=date.today()):
        if type(startdate) is str:
            startdate = date.fromisoformat(startdate)
        if type(enddate) is str:
            enddate = date.fromisoformat(enddate)

        self.cur.execute("SELECT * FROM housecharges WHERE chargedate BETWEEN ? AND ?", (startdate, enddate, ))
        rows = self.cur.fetchall()
        # todo return sums?
        return rows

    # def fetch_one(self, name, startdate, enddate):
        self.cur.execute("SELECT * FROM housecharges WHERE employee=? AND chargedate BETWEEN ? AND ?", (name, startdate, enddate))
        rows = self.cur.fetchall()
        return rows

    def insert_tournament(self, result_table, submit_date):
        """Ingests a list of lists containing the results from a single tournament.
           Adds results to golf_results.db        

        Args:
            result_table (list[list]): list of lists containing results from a single tournament
            submit_date (date): today's date, when database is updated.  Used for season changes 
        """
        print(submit_date)
        # Append the submission date to result_table...this is done in-place
        [i.append(submit_date) for i in result_table]
        
        for i in result_table:
            print([type(j) for j in i])
            
        self.cur.executemany(f"INSERT INTO golfresults VALUES (NULL, ?, ?, ?, ?, ?, ?)", result_table)
        self.conn.commit()
        

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
        rankings_sheet = wb['Rankings']
        
        result_table = []
        result_counter = 1
        result_with_ties = 0
        
        tournament = filename.split("/")[-1].split('.')[0]
        
        for entry in range(4, 100):
            if len(result_table) and rankings_sheet[f"B{entry}"].value:
                if (rankings_sheet[f"B{entry}"].value > rankings_sheet[f"B{entry-1}"].value):
                    result_counter += 1
                    result_with_ties += 1
                    if len(result_table)+1 > result_counter:
                        result_counter = result_with_ties
                else:
                    result_with_ties += 1
            else:
                result_with_ties += 1
            
            # Remove last (blank) value added by Excel Pivot Table
            if rankings_sheet[f"B{entry}"].value:
                if rankings_sheet[f"B{entry}"].value > 9999:
                    continue
            
            # Remove blank rows
            if rankings_sheet[f"C{entry}"].value in [None]:
                continue
            

                
            # Indicate ranking of "0" for teams that were CUT
            if rankings_sheet[f"B{entry}"].value == 9999:
                tournament_result = 0
            else:
                tournament_result = result_counter
                
            #player = [tournament, name, final score, ranking]
            player = [tournament, rankings_sheet[f"A{entry}"].value, rankings_sheet[f"B{entry}"].value, tournament_result, result_with_ties]
            result_table.append(player)
        
        # print(result_table)

        wb.close()
        # print((filename.split('.')[0].split('\\')[-1]))
        
        self.insert_tournament(result_table, date.today())

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
    folder = "./test_results/"
    # Fix for no network access...unkown reason why...
    # folder = "Z:/"

    files = os.listdir(folder)
    
    golf_files = [i for i in files if ".xlsm" in i]
    # print(files)
    # print(golf_files)
    
    start = input('Type ENTER to import all files listed above\n')
    
    for excel in golf_files:
        print(excel)
        db.manual_entry(f'{folder}{excel}')
            