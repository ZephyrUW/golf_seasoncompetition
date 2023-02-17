# Virtual Environment
# source my_env_project/bin/activate
# https://linuxopsys.com/topics/create-python-virtual-environment-on-ubuntu

# Scrape prior tournaments (based on date??) to get 1-3 place finishers.
# Track throughout season.  Create race?? graph.
# Email results

# Bump Plot https://towardsdatascience.com/7-visualizations-with-python-to-express-changes-in-rank-over-time-71c1f11d7e4b
# https://github.com/kartikay-bagla/bump-plot-python/blob/master/bumpplot.py

import os
import sqlite3
import openpyxl
import pandas

from rich import print, traceback
traceback.install()

def results_files(search_location):
    xlsm_files = [file for file in os.listdir(search_location) if file[-5:] == ".xlsm"]
    print(xlsm_files)
    result_table_full = []
         
    for xl_file in xlsm_files:
        tournament = xl_file.split(".")[0]
        wb = openpyxl.load_workbook(f"{search_location}{xl_file}", read_only=True, keep_vba=True, data_only=True)
        rankings_sheet = wb["Rankings"]
        
        result_table = []
        result_counter = 1
        result_with_ties = 0
        
        for entry in range(4, 70):
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
                tournament_result = "0"
            else:
                tournament_result = result_counter
                
            #player = [tournament, name, final score, ranking]
            player = [tournament, rankings_sheet[f"A{entry}"].value, rankings_sheet[f"B{entry}"].value, tournament_result, result_with_ties]
            result_table.append(player)
        result_table_full.append(result_table)

        # need to add result tabel to full results
    print(result_table_full)
    
def save_to_db(results):
    pass
    
# ?? Save to db
# Create Bump Plot
# Email Results
# Export HTML?

if __name__ == "__main__":
    results_files(search_location="./test_results/")
    
    