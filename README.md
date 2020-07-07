# py-excel-parser

### [Description in Russian](./description_rus.md)
### CLI utility for parsing excel file and form output json
 - #### Used `python 3.8`
 - #### Support `.xlsx`, `xlsm`, `xltx`, `xltm` Excel extensions

### Prerequisites:
  - #### Execute `git clone` of current repo
  - #### Create virtual `environment` by executing `python3.8 -m venv dev-env`
  - #### Activate `dev-env` by executing `source dev-env/bin/activate`
  - #### Install necessary dependencies by `pip install -r requirements.txt`

### Launch:
  - #### CLI keys description:
    - ##### `--excel_path`: Path to excel file [required]
      - Example: `python main.py --excel_path /user/income_excel.xlsx`
    - ##### `--json_path`: Path to output json file. By default named as  input excel file and put the same path.

  - #### Complete command example:
    - `python main.py --excel_path /home/ik/workspace/data.xlsm --json_path  /home/ik/output.json`

  - #### While executing script you will get messages:
    - About checking output file existence
    - Performing loading of workbook
    - Which sheet was parsed
    - Parsing procedure completion
