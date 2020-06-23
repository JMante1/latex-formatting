# -*- coding: utf-8 -*-
"""
Created on Tue Jun 23 12:34:51 2020

@author: JVM
"""
import os
import pandas as pd

def table_to_latex(path_in, path_out, sheet_name = "Sheet1",
                   table_label = "table1", margin = 2.5, page_width = 21,
                   excel = True, header_col = True, header_row = True,
                   bold_final = True, format_string= '{:,.0f}'):
    """
    Reads in excel or csv tables and converts them to text that can be pasted
    into latex to show a table.
    

    Parameters
    ----------
    path_in : STRING
        The full file path at which the input excel or csv file is found
        For example 'C://User//excel_file.xlsx'
    path_out : STRING
        The full file path at which the resulting textfile should be saved
        For example 'C://User//output.txt'
    sheet_name : STRING, optional, default "Sheet1".
        The name of the sheet to look at if an excel file is being read in
    table_label : STRING, default "table1".
        The name for the figure reference incorporated. Will allow latex
        to cite using \cite{table_label}
    margin : FLOAT, default 2.5
        The width in centimeters from the page edge to the right edge of the
        table/text
    page_width : FLOAT, default 21
        The width in centimeters of the page. The default is for A4.
    excel : BOOL, default True
        Whether the file being read in is xlsx. If False a csv will be read in
    header_col : BOOL, default True
        If True there is a double line between the first column and the rest
        otherwise there is just a single line
    header_row : BOOL, default True
        If true there is a double line between the first row and the further
        rows. If false there is a single line.
    bold_final : BOOL, default True
        If true the final row will have bold text
    format_string : STRING, default '{:,.0f}'
        The pyformat string to format all number columns in the table. For
        more information see the Number section at: https://pyformat.info/

    Returns
    -------
    None.
    
    Requires
    --------
    import pandas as pd
    
    Example
    -------
    file_name = "test"
    cwd = os.path.dirname(os.path.abspath("__file__")) #get current working directory
    path_in = os.path.join(cwd,f"{file_name}.xlsx")
    path_out = os.path.join(cwd,f"output.txt")
    
    table_to_latex(path_in, path_out)

    """
    rows = []
    
    #read in data, if header row it needs to be read in separately to prevent the column being
    #seen as not fully numeric. The read in of either csv or excel happens.
    if header_row:
        if excel:
            df = pd.read_excel (path_in, sheet_name = sheet_name, header=None, skiprows=1)
            header = pd.read_excel (path_in, sheet_name = sheet_name, header=None, nrows = 1).iloc[0]
        else:
            df = pd.read_csv (path_in, header=None, skiprows=1)
            df = df.dropna(how = 'all')
            header = pd.read_csv (path_in, header=None, nrows = 1).iloc[0]
        
        #if heder row is read make sure all column names are strings and then put
        # '&' between each column header. Finish the header with the line end and
        #two lines
        row_string = header.tolist()
        row_string = [str(i) for i in row_string]
        row_string = ' & '.join(row_string)
        row_string += r' \\ \hline  \hline'
    else:
        if excel:
            df = pd.read_excel (path_in, sheet_name = sheet_name, header=None)
        else:
            df = pd.read_csv (path_in, header=None)
            df = df.dropna()
    
           
    
    
    #create the start text with column width formatting
    width = (page_width - margin*2 )/ len(df.columns)
    begin_string =   r'\begin{table*}[ht]'+'\n\t'+r'\caption{caption goes here}'+'\n\t'+r'\begin{tabular}{|'
    
    #if there is a header_column it will be separated from the rest with two '|'
    for col in range(0,len(df.columns)):
        if col == 0 and header_col:
            begin_string += f"p\u007b{width:.1f}cm\u007d||"
        else:
           begin_string += f"p\u007b{width:.1f}cm\u007d|" 
    begin_string += '} \hline'
    
    #add the beginning string to the table strings
    rows.append(begin_string)
    
    #if the header_row exists add the header row string
    if header_row:
        rows.append(row_string)
    
    
    
    #ensure column names are strings as that is needed to map the formatting of numeric columns
    df.columns = [str(i) for i in list(range(0,df.shape[1]))]
    
    #format numeric columns based on the format_string
    numerics = ['int16', 'int32', 'int64', 'float16', 'float32', 'float64']
    numerics_df = df.select_dtypes(include=numerics)        
    for col_name in numerics_df.columns:
        df[col_name] = df[col_name].map(format_string.format)
    
    #chain columns with '&" and add bold formatting to the final row if bold_final = True
    for row_index, row in df.iterrows():
        row_string = row.tolist()
        
        if bold_final and row_index == len(df)-1:
            #remove nan in the final row
            for index, cell in enumerate(row_string):
                if cell=='nan':
                    row_string[index] = ''
                    
            row_string = [ r'\textbf{'+cell+r'}' for cell in row_string]
            
        row_string = ' & '.join(row_string)
        row_string += r' \\ \hline'
        rows.append(row_string)
    
    #add the end row section including the 'figure label'    
    rows.append('\end{tabular}\n\t\label{tab:'+table_label+'}\n\end{table*}')
    
    # put an enter and tab between the header section, every internal row,
    #and the end row
    rows = '\n\t'.join(rows)
    
    #print out the final table to the document at path_out
    with open(path_out, "w") as text_file:
        text_file.write(rows)
        
    return

file_name = "Pivot"
cwd = os.path.dirname(os.path.abspath("__file__")) #get current working directory
path_in = os.path.join(cwd,f"{file_name}.csv")
path_out = os.path.join(cwd,f"output.txt")

table_to_latex(path_in, path_out, excel = False)
