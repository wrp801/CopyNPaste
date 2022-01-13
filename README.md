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


## Clone the Repository 

In order to do this, you will need **git**, if that is not installed on your computer you can download it [here](https://git-scm.com/downloads).

To download the module, in Git Bash or Anaconda Powershell, run the following:

```bash 
cd .xlwings
git clone https://github.com/wrp801/CopyNPaste.git
```
This will create a new directory in the `.xlwings` directory.

## On initial Installation
Now that the config file is generated, run the following script in the *Anaconda Powershell*:
```powershell
python .xlwings\CopyNPaste\set_config.py 
```

## For updates 
If you have re-pulled from the repository, then will need to redo the config setup. That can be done by deleting the initial file, generating it again, then running the python script once more. Run the following to do so: 

```powershell
rm -fr .xlwings\xlwings.conf
xlwings config generate
python .xlwings\CopyNPaste\set_config.py 

```

This will automatically configure the settings to import the user defined functions (UDF's) into the Excel Workbook. 

## Excel Setup
The next thing required is to properly configure Excel. It is important to note that these UDF's will only work on a macro enabled workbook, so **ensure that the workbook is saved as a .xlsm**. 

The next steps need to only be applied one time.

1.  Go to File > Options > Trust Center > Trust Center Settings > Macro Settings 
	- Check the `Trust access to the VBA project object model` option
2. In the Developer tab on the Excel Ribbon, select the `Visual Basic` option
	- Go to Tools > References and check the `xlwings` box. 


## Use 
Now, whenever you have a .xlsm workbook, you simply need to select the **Import Functions** option on the `xlwings` ribbon to import the function. It should take a few seconds, but after that you will be able to see the functions appear once you hit the `=` in a cell. 

## Functions 

##### copy_and_paste
This function receives a variable number of arguments. 

