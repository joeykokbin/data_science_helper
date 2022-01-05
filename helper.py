import os
import pandas as pd
import xlwings as xw

def read_pswd_excel(link, sheet, excel_range = "", return_df = True):
    
    ''' 
    Simple helper function to read password-protected Excel files
    
    link:           Link of excel file to read.
    sheet:          Sheet of Excel File. (will throw error if sheet is not found)
    excel_range:    string of an excel range such as "A1:G20"
    return_df:      To return as list of list (False) or a dataframe (True).
    
    Each element in the list of list is a row. 
    Each element in a row would be a column, starting from excel_range
    
    '''
    
    if not os.path.exists(link):
        print("link does not exist. Please try again.")
        raise Exception("LinkNotFoundError")
        
    app = xw.App()
    filebook = xw.Book(link)
    
    data = filebook.sheets[sheet].range(excel_range).value
    
    filebook.close()
    app.quit()
    
    if return_df:
        df = pd.DataFrame(data[1:], columns = data[0]).dropna(how = "all", axis = "rows").dropna(how = "all", axis = "columns")
        df.columns = map(str.upper, df.columns)

    else:
        df = data
    
    return df