"""
Module imports an Excel workbook and outputs a YAML file 
"""

import sys
import logging
import argparse

from openpyxl import load_workbook
from yaml import load, dump

LOGLEVEL = logging.DEBUG

# Create logger for module
logger = logging.getLogger(__name__)
logger.setLevel(LOGLEVEL)

# Create logging formatter
formatter = \
    logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

# Create a console handler and add it to logger
ch = logging.StreamHandler()
ch.setLevel(LOGLEVEL)
ch.setFormatter(formatter)
logger.addHandler(ch)

class Sheet():
    """
    Parses an Excel worksheet and builds Python objects based on the
    following rules:
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

    """
    
    def __init__(self, worksheet):
        '''
        Initialize the worksheet object
        '''
        
        self._ws = worksheet
        self._objects = []
        
        self.parse()
        
    def parse(self):
        '''
        Parse the worksheet and identify objects
        '''
        
        logger.debug('Beginning parse() of %s' , self._ws)
        
        this_obj = []
        for row in self._ws.rows:
            this_row = []
            empty = True
            for cell in row:
                if cell.value:
                    this_row.append(str(cell.value))
                    empty = False
                else:
                    this_row.append(cell.value)
            if not empty:
                this_obj.append(this_row)
            else:
                self._objects.append(this_obj)
                this_obj = []
        if this_obj:
            self._objects.append(this_obj)
        
        logger.debug('Completed parse() of %s', self._ws)
            
    def getCollections(self):
        '''
        Return the collection of objects contained within the sheet
        '''
        
        return self._objects
    
class Collection():
    '''
    Individual objects to be converted into other formats
    '''
    
    def __init__(self, array):
        
        self._this_collection = self.parseObject(array)
        
    def clean(self, array, side='left'):
        result = list(array)
        if side == 'both':
            result = self.clean(array, 'right')
        del_count = 0
        if side == 'right':
            result.reverse()
        for idx in range(len(result)):
            if result[idx] == None:
                del_count += 1
            else:
                break
        if del_count :
            del result[:del_count]
        if side == 'right':
            result.reverse()
        return result
    
    def buildParentList(self, indentList):
        result = []
        parentQueue = []
        
        for idx in range(len(indentList)):
            
            if indentList[idx] == 0:
                result.append(None)
                parentQueue = []
            elif indentList[idx] > indentList[idx - 1]:
                parentQueue.append(idx - 1)
                result.append(parentQueue[-1])
            elif indentList[idx] == indentList[idx - 1]:
                result.append(parentQueue[-1])
            elif indentList[idx] < indentList[idx - 1]:
                popNumber = indentList[idx - 1] - indentList[idx]
                for i in range(popNumber):
                    parentQueue.pop()
                result.append(parentQueue[-1])
                
        return result

    def buildObject(self, array, final=False):
        ''' Build an input 2-dimensional array into a Python object '''
        
        # Evaluate base case of a single element / single line
        if (len(array) == 1) or (final == True):
            initObject = array[0]
            var = initObject[0]
            if len(initObject) > 2:
                # This is a list
                val = list(initObject[1:])
                return { var : val }
            elif len(initObject) == 2:
                # This is a single key-value pair
                val = initObject[1]
                return { var : val }
            else:
                # This is a single element and should return the element
                return initObject[0]
            
        # Evaluate case of multi-line/multi-dimensional array
        elif len(array) > 1:
            header = array[0]
            var = header[0]
            temp_list = []
            if len(header) == 1:
                # Evaluate case of multi-line list; return single-element list
                for item in array[1:]:
                    temp_list.append(self.buildObject([item]))
                return { var : temp_list }
            elif len(header) > 1:
                # Evaluate case of multi-line dict; return single-element list
                keys = header[1:]
                for item in array[1:]:
                    temp_dict = {}
                    for idx in range(len(keys)):
                        key = keys[idx]
                        current_obj = self.buildObject([[key , item[idx]]])
                        temp_dict[key] = current_obj[key]
                    temp_list.append(temp_dict)
                return { var : temp_list }
                
    def parseObject(self, array):
        '''
        Take raw input array, clean it, use buildObject to 
        compile a top-level Python object out of array
        '''
        
        # Clean right side of array; remove extra empty (None) cells
        initObject = [None for row in range(len(array))]
        for idx in range(len(array)):
            initObject[idx] = self.clean(array[idx], side='right')
        ### print "initial clean list is ", initObject
        
        # Clean left side of array; build list of line indentions
        indentList = [None for row in range(len(initObject))]
        for idx in range(len(initObject)):
            temp_item = self.clean(initObject[idx], side='left')
            indent = int(len(initObject[idx]) - len(temp_item))
            indentList[idx] = indent
            if indent > 0:
                initObject[idx] = temp_item
        ### print "final clean list is ", initObject
        ### print "indent list is", indentList
        
        workingResult = list(initObject)
        ### print "working result is ", workingResult
        
        parentList = self.buildParentList(indentList)
        ### print "parent list is ", parentList
        
        maxIndent = max(indentList)
        ### print "max indent is ", maxIndent
        
        # Compile all elements down to indent level 0
        while maxIndent > 0 :
            logger.debug('evaluating while maxIndent > 0...')
            logger.debug('maxIndent = %d', maxIndent)
            logger.debug('workingResult = %s', workingResult)
            logger.debug('indentList: %s', indentList)
            logger.debug('parentList: %s', parentList)
            for idx in reversed(xrange(len(workingResult))):
                logger.debug('evaluating idx: %s', str(idx))
                if indentList[idx] == maxIndent:

                    this_parent = parentList[idx]
                    logger.debug('this parent: %s', this_parent)
                    line_count = idx - this_parent
                    logger.debug('line count: %s' , str(line_count))
                    
                    this_obj = \
                        workingResult[this_parent:(this_parent+line_count) + 1]
                    logger.debug('this_obj: %s', str(this_obj))

                    logger.debug('building temp_item...')
                    temp_item = self.buildObject(this_obj)
                    logger.debug('temp_item = %s', temp_item)

                    workingResult[this_parent] = [temp_item]
                    logger.debug('working result[%s] = %s', \
                        this_parent, workingResult[this_parent])
                    
                    start_del = this_parent + 1
                    end_del = this_parent + line_count + 1
                    logger.debug('deleting workingResult[%d:%d]...', \
                        start_del, end_del)
                    del workingResult[start_del : end_del]
                    logger.debug('deleting indentList[%d:%d]...', \
                        start_del, end_del)
                    del indentList[start_del : end_del]
                    logger.debug('recalculating parentList...')
                    parentList = self.buildParentList(indentList)
                    
                    break
            logger.debug('reseting maxIndent')
            maxIndent = max(indentList)
        
        ### print "level 0 working result: \n", workingResult
        
        for idx in range(len(workingResult)):
            this_obj = [workingResult[idx]]
            ### print "this obj = " , this_obj
            temp_item = self.buildObject(this_obj, final=True)
            workingResult[idx] = temp_item
        
        ### print "working result: \n ", workingResult
        
        ### print "\n\n\n\n"
        ### print "Items:"
        final_dict = {}
        for item in workingResult:
            ### print item
            if type(item) == dict:
                for key in item.keys():
                    final_dict[key] = item[key]
        ### print "final dict: ", final_dict

        return [final_dict]

    def buildOutput(self, type='yaml'):
        '''
        Generates the text version of the object, for use in
        writing to output files
        '''
        
        if type == 'yaml':
            result = dump(self._this_collection, default_flow_style=False)
        
        return result

    @property
    def collection(self):
        return self._this_collection
    
    
    
def xlyaml(source, output=None, format='yaml'):
    '''
    Primary function for building YAML document from workbook
    '''
    
    logger.info('Opening workbook %s', source)
    
    # Open workbook
    try:
        wb = load_workbook(source)
    except:
        print "Unable to open workbook:", source
        print "Please check filename and try again."
        print "\n\n"
        sys.exit(2)
    
    # Gather names of worksheets within the workbook
    worksheets = []
    for sheet in wb:
        worksheets.append(sheet)
        logger.debug('Found sheet: %s', sheet)
    logger.info('Found %d sheets', len(worksheets))    
    
    for sheet in worksheets:
        logger.info('Beginning evaluation of sheet %s', sheet)
        sheetObject = Sheet(sheet)
        collectionObjects = sheetObject.getCollections()
        
        # Set up output file that will correlate to the current sheet
        sheetName = str(sheet.title)
        outFile = open(sheetName + '.yml', 'w')
        logger.debug('Opened output file %s', outFile)
        outFile.write('# ' + sheetName)
        outFile.write('\n')
        outFile.write('---')
        outFile.write('\n')
        
        # Cycle through collection objects contained in sheet
        logger.debug('Building collection objects found in %s', sheet)
        for collection in collectionObjects:
            this_obj = Collection(collection)
            outFile.write(this_obj.buildOutput(type=format))
            outFile.write('\n')
        
        # Clean up current output file before moving to next sheet
        outFile.write('...')
        outFile.close()
        logger.info('Completed output file %s', outFile)
        
    logger.info('Completed execution of xlyaml')
    
if __name__ == "__main__":
    """
    This runs if the program is called from the command line.
    """

    parser = argparse.ArgumentParser()
    parser.add_argument("source", type=str, 
            help="Source Excel file")
    parser.add_argument("-l", "--loglevel", 
            choices=['INFO', 'DEBUG'], default = 'INFO',
            help="Logging level; defaults to 'INFO'")
    args = parser.parse_args()
    
    if args.loglevel == 'INFO':
        logger.setLevel(logging.INFO)
        ch.setLevel(logging.INFO)
    elif args.loglevel == 'DEBUG':
        logger.setLevel(logging.DEBUG)
        ch.setLevel(logging.DEBUG)
        logger.debug('Setting logging level to DEBUG')
        
    source = args.source
    xlyaml(source)
    
'''####### TEST CASES #########

xlyaml('sample.xlsx')'''