# copy_n_paste 

This provides a simple script for allowing a way to copy and paste a range of cells. The implementation is in Python, but no coding is required for end users. Follow the setup instructions outlined below and you will be able to use this on all Excel versions of 2019 and above. It is also worth noting that this tool is only guaranteed to work on Windows. 


## Setup 
First thing to do is to install Python on your machine (if it is not already installed). For simplicity, the Anaconda distribution is recommended (found [here](https://docs.conda.io/en/latest/miniconda.html)). Select the windows installer, then upon completion open the anaconda powershell or anaconda command prompt and run the following:

```bash
conda install pandas 
conda install xlwings
```
This will install the two dependencies needed for the script. Upon completition run the following:

```bash
xlwings addin install
```

This will create a .xlwings directory in your home directory (`C:\Users\<username>\.xlwings`). 


