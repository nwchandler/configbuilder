# configbuilder
Configuration Build Automation Utilities

Primary utilities are xlyaml and textbuilder. xlyaml provides the ability to convert properly-formatted Excel workbooks and worksheets into YAML files. textbuilder provides the ability to convert Jinja2 formatted template files into final documents, substituting variables from the YAML files produced by xlyaml.

## Using xlyaml

In order to use xlyaml, it is important to ensure that Excel documents follow the proper formatting. xlyaml parses an Excel worksheet and builds Python objects based on the following rules:

    - Worksheet contains some number of similar object definitions
    - Top line in worksheet is the first object; subsequent objects are
    separated by a blank row
    - Lines not separated from previous object by blank row are part
    of the previous object
    - Two cells that are side-by-side represent key-value pairs
                _________________________
                |    Key    |   Value   |
    - Lists are represented by a left-cell, blank right-cell, subsequent
    single cells in a column under the blank right-cell
                ________________________
                |  Name    |   (blank)  |
                ________________________
                | (blank)   |   Item 1  |
                ________________________
                | (blank)   |   Item 2  |
                ________________________
                | (blank)   |   Item 3  |
     - Dictionaries are similar to lists, but with named indexes
     in place of blank cells; supports multiple named indexes
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

	usage: xlyaml.py [-h] [-l {INFO,DEBUG}] source

	positional arguments:
	  source                Source Excel file

	optional arguments:
	  -h, --help            show this help message and exit
	  -l {INFO,DEBUG}, --loglevel {INFO,DEBUG}
	                        Logging level; defaults to 'INFO'

## Using textbuilder

To do...

## Some notes on running Python files...

Windows executable files can be obtained by [contacting Nick](mailto:nchandler@burwood.com); if you are not familiar with Python, this is the preferred way to use the configbuilder tools currently. If you prefer to leverage your own Python interpreter, I recommend that you look at using virtualenv and virtualenvwrapper in order to have sandbox environments for various Python projects you might want to use. (Nick is happy to help you with this process, but be prepared to do a little googling... Familiarity with Python will go a long way here.)

Once you have your virtual environment running (or not, if you don't want to), use the command 'pip install -r requirements.txt' in order to install the extra modules that xlyaml and textbulider use.  

## FAQ

***How do I report bugs?***

Use the [Issues][issues] feature in GitHub to report bugs.


***How do I request new features?***

Use the [Issues][issues] feature in GitHub to create a new feature request.

## Contributing to configbuilder

configbuilder is written in Python and is currently maintained by Nick Chandler. If you would like to work on the code, please [contact Nick](mailto:nchandler@burwood.com). If you would like to contribute, but don't feel comfortable coding, documentation contributions are welcome. Also, please let us know if you would like to see any features added, using the [Issues][issues] page, so we can continue to improve the tools and make them more useful over time.
[issues]: https://github.com/Burwood/configbuilder/issues