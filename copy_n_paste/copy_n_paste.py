import xlwings as xw
import pandas as pd 

@xw.func 
@xw.ret(expand = 'table')
@xw.ret(index = False,header = False)
def copy_and_paste(n_col,*args):
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
        

