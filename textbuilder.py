''' Builds text files from templates and YAML input '''

import os
import logging
import sys
import string
import random
import argparse

import yaml
import jinja2

logging.basicConfig(format='%(asctime)s - %(levelname)s - %(message)s', level=logging.INFO)

def id_generator(size=12, chars=string.ascii_uppercase + string.digits):
    """
    Generates a random string, default size 12 characters
    Thanks to Stack Overflow user Ignacio Vazquez-Abrams for this function
    Search for "Random string generation with upper case letters
    and digits in Python" on Stack Overflow
    """
    return ''.join(random.choice(chars) for _ in range(size))

def textbuilder(temp=None, varfile=None, outfile='results.txt', fileid=None):
    '''
    Returns string containing text of templates
    Optionally outputs to file 
    
    :param temp: Template file
    :type temp: File, formatted for jinja2 processing
    
    :param varfile: File containing variables referenced in temp
    :type varfile: File, formatted in YAML
    
    :param outfile: Output filename
    :type outfile: String, containing filename
    
    :param fileid: If writing to multiple files, this tag distinguishes
            the files from one another. Must be a valid tag
            found in the varfile. Should only be used if the
            top line of the template is a loop that will generate
            multiple similar segments of text. For example,
            if the top level is {% for host in hosts %}, the
            fileid might be something like host.name or 
            host.id .
    :type fileid: String
    
    :rtype: string
    '''
    
    logging.debug('Running textbuilder...')
    result = None
    
    # If fileid is set, then outfile will not be used
    if fileid:
        outfile = None
    
    # Generate a cookie to identify where to separate docs if the function
    # needs to generate multiple output files
    if fileid:
        logging.info('Generating random ID for multi-file output')
        random_id = id_generator()
        logging.debug('Random ID is %s', random_id)
        logging.info('Creating cookie for multi-file output')
        cookie = random_id + '{{ ' + fileid + ' }}' + random_id
        logging.debug('Cookie is %s', cookie)
    
    # Open template file
    try:
        tempfile = open(temp, 'r')
        logging.info('Opened template: %s', temp)
        # Insert cookie into template if multiple output files needed
        if fileid:
            splitter = []
            # Insert the first line of tempfile
            splitter.append(tempfile.readline())
            # Insert the cookie
            splitter.append(cookie)
            # Insert the remainder of tempfile
            splitter.append(''.join(tempfile.readlines()))
            template = ''.join(splitter)
        else:
            template = tempfile.read()
        logging.debug('Template contents:\n%s', template)
    except:
        logging.error('Failed to open template: %s', temp)
        print "\n"
        print "Failed to open template:", temp
        print "Please check filename and try again."
        print "\n"
        raise
        
    # Open variable file and load variables
    try:
        vars = yaml.load(open(varfile, 'r'), Loader=yaml.Loader)
        logging.info('Opened variable file: %s', varfile)
        logging.debug('Variable file contents:\n%s', vars)
    except:
        logging.error('Failed to open variable file: %s', varfile)
        print "\n"
        print "Failed to open variable file:", varfile
        print "Please check filename and try again."
        print "\n"
        raise    
    
    # Set up variable reference, which should match what users will put
    # in their template files
    base = os.path.basename(varfile)
    varref = os.path.splitext(base)[0]
    logging.info('Variable reference is %s', varref)
    
    the_template = jinja2.Template(template)
    # Jinja2 doesn't allow keywords to be variables, so we have to build
    # a string which will subsequently be evaluated in order to keep this
    # as general as desired.
    render_string = 'the_template.render('+varref+'=vars)'
    rendered = eval(render_string)
    logging.debug('rendered: %s', str(rendered))
    
    # Write output file(s)
    # Write multiple files if fileid is set
    if fileid:
        # Split rendered string so we can strip the cookie out
        rendered_split = rendered.split(random_id)
        result_split = []
        # The first element may be a newline left over from initial for loop
        if rendered_split[0] == '\n':
            del(rendered_split[0])
        # Names should be even list items, and data should be odd
        for idname in range(0, len(rendered_split), 2):
            # Assume output extension is .txt for now
            fname = rendered_split[idname] + '.txt'
            logging.info('Writing file %s', fname)
            with open(fname, 'w') as f:
                f.write(rendered_split[idname + 1])
            result_split.append(rendered_split[idname + 1])
        # Prepare a single result, w/o cookies, to return
        result = ''.join(result_split)
    else:
        result = rendered
        try:
            ofile = open(outfile, 'w')
            logging.info('Writing results to %s',ofile)
            ofile.write(result)
            ofile.close()
            logging.info('Closed output file %s',ofile)
        except:
            pass
        
    tempfile.close()
    logging.info('Closed template: %s', temp)
    
    return result

if __name__ == "__main__":
    """
    This runs if the program is called from the command line.
    """

    parser = argparse.ArgumentParser()
    parser.add_argument("template", type=str, 
            help="Template file, format=Jinja2")
    parser.add_argument("varfile", type=str, 
            help="File containing variables")
    parser.add_argument("-o", "--outfile", type=str,
            help="""If outputting to a single file, this is the filename.
            If omitted, defaults to 'results.txt'""")
    parser.add_argument("-i", "--fileid", type=str,
            help="""If output to multi-file, this is the tag used to
            identify those files. Must refer to tag in variable file.""")
    args = parser.parse_args()
    
    outfile = 'results.txt'
    fileid = None
    template = args.template
    varfile = args.varfile
    if args.outfile:
        outfile = args.outfile
    if args.fileid:
        fileid = args.fileid
    
    logging.debug("""Running textbuilder with options:
            template: %s
            varfile: %s
            outfile: %s
            fileid: %s""",
            template, varfile, str(outfile), str(fileid))

    textbuilder(template, varfile, outfile, fileid)


########## TEST CASES ############

'''test = textbuilder(temp='configtemp.txt', varfile='switches.yml', outfile='results.txt', fileid='host.name')
print test'''

'''vm = textbuilder(temp='vmtemplate.txt', varfile='vmhosts.yml', outfile='results.py')
print vm'''
