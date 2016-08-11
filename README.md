# configbuilder
Configuration Build Automation Utilities

Primary utilities are xlyaml and textbuilder. xlyaml provides the ability to convert properly-formatted Excel workbooks and worksheets into YAML files. textbuilder provides the ability to convert Jinja2 formatted template files into final documents, substituting variables from the YAML files produced by xlyaml.

## Using xlyaml

In order to use xlyaml, it is important to ensure that Excel documents follow the proper formatting. Two formats are accepted for sheets.

    The first is a simple table format, such as:
        __________________________________________________
        |    Key1     |   Key2    |   Key3    |   Key4    |  <-- Header Row
        __________________________________________________
        |    Item1    |   Item 2  |   Item 3  |   Item 4  |  <-- Object 1
        __________________________________________________
        |    Item1    |   Item 2  |   Item 3  |   Item 4  |  <-- Object 2

    
    The second format is a list of objects. This format is better suited
    for scenarios where objects may have different numbers of variables
    or need to support list or dictionary attributes. The following rules
    apply:
        - Worksheet contains some number of similar object definitions
        - Top line in worksheet starts the first object; subsequent objects are
        separated by single blank rows
        - Lines not separated from previous object by blank row are part
        of the previous object
        - Two cells that are side-by-side represent key-value pairs; e.g.:
                _________________________
                |    Key    |   Value   |
        - Lists are represented by a left-cell, blank right-cell, subsequent
        single cells in a column under the blank right-cell; e.g.:
                ________________________
                |  Name    |   (blank)  |
                ________________________
                | (blank)   |   Item 1  |
                ________________________
                | (blank)   |   Item 2  |
                ________________________
                | (blank)   |   Item 3  |
         - Dictionaries are similar to lists, but with named indexes
         in place of blank cells; supports multiple named indexes; e.g.:
                __________________________________________________
                |    Name    |   ID1     |    ID2     |   ID3     |
                __________________________________________________
                |    (blank)  |   Item 1  |    Item 1 |   Item 1  |
                __________________________________________________
                |    (blank)  |   Item 2  |    Item 2 |   Item 2  |
                __________________________________________________
                |    (blank)  |   Item 3  |    Item 3 |   Item 3  |

If the workbook that you want to convert has more than one worksheet, each sheet will be parsed. The results of each sheet will be placed in separate files, named after the worksheets. For example, if you have three sheets: 'sheet1', 'sheet2', and 'sheet3', you will obtain three files: 'sheet1.yml', 'sheet2.yml', and 'sheet3.yml'. Because the YAML file name is later used for variable naming purposes in textbuilder, you should name sheets in a clear manner.

Once you have a properly formatted workbook, you can execute xlyaml with the following usage parameters; note, the only required option is a source Excel file:

	usage: xlyaml.py [-h] [-l {INFO,DEBUG}] [-f {table,list}] source

    positional arguments:
      source                Source Excel file

    optional arguments:
      -h, --help            show this help message and exit
      -l {INFO,DEBUG}, --loglevel {INFO,DEBUG}
                            Logging level; defaults to 'INFO'
      -f {table,list}, --sourceformat {table,list}
                            Format of Excel spreadsheet

## Using textbuilder

To do...