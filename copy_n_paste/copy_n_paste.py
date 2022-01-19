import xlwings as xw
import pandas as pd 
import numpy as np

@xw.func 
@xw.ret(expand = 'table')
@xw.ret(index = False,header = False)
def create_array(n_col,*args):
    """
    This function is designed to copy and paste a range of cells (designated by *args) 
    n times (designated by the n_col argument). This performs a row wise operation where the 
    ith row of args will be copied the number of times as specified by the ith row of n_col. 

    Example: 
    copy_and_paste(A1:A5,B1:B5,D1:D5)
    This would return the two columns (B and D) where the number of rows will be equal to the sum of A1:A5 (i.e the number of times to be copied)
    """
    ## iterate through the number of columns to copy
    df = pd.DataFrame()
    col_num = 0
    for arg in args:
        ret_list = []
        for a,n in zip(arg,n_col):
            col_name = "x" + str(col_num) ## temporary column name, will be dropped on return
            for i in range(int(n)):
                ret_list.append(a)
            col_num+= 1
        temp_df = pd.DataFrame({col_name: ret_list})
    
        df = pd.concat([df,temp_df],axis= 1)
    return df 

@xw.func 
@xw.ret(expand = 'table',index = False,header = False)
def create_array_continuous(n_col,cell_range):
    df = pd.DataFrame()
    nrows = len(cell_range)
    ncols = len(cell_range[0])

    for n,row in zip(n_col,cell_range):
        n_val = int(n)
        for i in range(n_val):
            row_list = []
            for col in row:
            ## This will go row by row, moving left to right through each cell in each row
                row_list.append(col)
            arr = np.array(row_list).reshape(1,ncols)
            temp_df = pd.DataFrame(arr)
            df = pd.concat([df,temp_df])
    return df 

        

