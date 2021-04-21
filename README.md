# Bulk-Email

This repository contains a simple python script that can send emails to multiple individuals. This code was built to primarily solve the problem of sending grades and scores of students individually. 

It is suggested to run the script with 

```bash
python email.py --test_mode True
```

to run it in test mode and not send any emails, but run through everything else. More details about test mode are detailed below.

## Usage with no attachments

The `Summary.xlsx`file with multiple sheets gives is for reference. It gives an idea of how to arrange information and sending emails to individuals whose email and names are present in `email_names.csv`. 

You can run the code by executing 

``` bash
python email.py
```

An example message body 

```
Hello AA,
 Please find attached a .txt file that summarises your grades and scores for the <INSERT COURSE HERE> course.
 It is suggested to open the document in landscape mode, if you are having trouble viewing it.

 Thank you

 Regards,
 Sawan Singh Mahara 
```

The script uses the details entered in `Summary.xlsx` and creates a `.txt ` file that will be attached to the email. It can be found in the `Text Files` folder.

## Usage with attachments and no xlsx file.

As an example, if you want to send a `PDF` to each individual, make sure that the `PDF` is named `<INDIVIDUAL EMAIL>.pdf` where for example if you want to send the email to `xyz@abc.com`, the pdf should be `xyz.pdf`. They must be put into the `Attachments `folder.

You can run the code by executing 

``` bash
python email.py --use_xlsx False --attachment_extention .pdf
```

The arguments `--use_xlsx` and `--attachment_extention` can be combined with other arguments as well. You can type `python email.py --help`  to know about the other supported arguments.

An example message body 

```
Hello AA,
 Please find attached a PDF for you.

 Thank you

 Regards,
 Sawan Singh Mahara 
```

The script uses the details entered in `Summary.xlsx` and creates a `.txt ` file that will be attached to the email. It can be found in the `Text Files` folder.

## Test Mode

If you want to run the script in a test mode not risking accidentally sending the wrong file to the wrong recipient, you can run the script as

```bash
python email.py --test_mode True
```

If you want to send an email but to some test emails, and not the intended recipients, run

``` bash
python email.py --test_mode_and_send True
```

This will prompt you to enter some test emails, that you don't mind sending it to, for example it can be your own email address again or trusted party's email address.

This will make the script perform the exact same way, but instead of sending the email to every recipient, it randomly selects an individual from the list of individuals in `email_names.csv` generates the necessary email body/attachments and sends their intended email to the test email addresses.

# email_names.csv

If you want to send emails to the individuals aaa@iitdh.ac.in, bbb@iitdh.ac.in,...., zzz@iitdh.ac.in, this `.csv` file is already setup.

If not, you can modify the `.csv` file or create your own in the same format and replace it here.

# Summary.xlsx

For the provided Summary.xlsx file, the corresponding `.txt` file for each individual is created and stored in `Text Files`. The contents of this folder can be deleted and your own `Summary.xlsx` file can be used as well. 

The script only works if there are exactly two sheets. If you need only one sheet or more than 2 sheets, you will have to modify the code.

The following lines in the code control the sheets.

```python
#%% XLSX Sheet Code
Email_names=np.loadtxt('email_names.csv',dtype=str,delimiter=',')
if use_xlsx:
    sheet1=pd.read_excel("Summary.xlsx",sheet_name=0)
    sheet2=pd.read_excel("Summary.xlsx",sheet_name=1)
    sheets=[sheet1,sheet2]
```

and

```python
for work in ["Sheet 1 Title", "Sheet 2 Title"]:
    a=sheets[itr].iloc[individual,2:] #Consider the cells from column C onward, row 'itr'
    a=a.to_string()+'\n \n \n'
```



### If there is only one sheet use

```python
#%% XLSX Sheet Code
Email_names=np.loadtxt('email_names.csv',dtype=str,delimiter=',')
if use_xlsx:
    sheet1=pd.read_excel("Summary.xlsx",sheet_name=0)
    sheets=[sheet1]  
```

and

```python
for work in ["Sheet 1 Title"]:
    a=sheets[itr].iloc[individual,2:] #Consider the cells from column C onward, row 'itr'
    a=a.to_string()+'\n \n \n'  
```



### If there are four sheets, use

```python
#%% XLSX Sheet Code
Email_names=np.loadtxt('email_names.csv',dtype=str,delimiter=',')
if use_xlsx:
    sheet1=pd.read_excel("Summary.xlsx",sheet_name=0)
    sheet2=pd.read_excel("Summary.xlsx",sheet_name=1)
    sheet3=pd.read_excel("Summary.xlsx",sheet_name=2)
    sheet4=pd.read_excel("Summary.xlsx",sheet_name=3)
    sheets=[sheet0,sheet1,sheet2,sheet3]
```

and

``` python
for work in ["Sheet 1 Title", "Sheet 2 Title","Sheet 3 Title", "Sheet 4 Title"]:
    a=sheets[itr].iloc[individual,2:] #Consider the cells from column C onward, row 'itr'
    a=a.to_string()+'\n \n \n'  
```

If there are different number of sheets, modify these lines accordingly. 

# Google Authentication

If the sender email ID is a google email ID, this script will work only if two-factor authentication is not enabled. 

If it is not enabled, then you have to head over to https://myaccount.google.com/lesssecureapps and enable less secure apps to use that email ID as the sender ID.

Once you have used the script, you can disable this to get back security.

