#############################################################################
#############################################################################
###                                                                       ###
###        pptx notes exporter v1.0 copyright Eric Jang 2012              ###
###        ericjang2004@gmail.com                                         ###
###                                                                       ###
#############################################################################
#############################################################################

#!/usr/bin/env python
import argparse, os, glob
from zipfile import ZipFile
from xml.dom.minidom import parse

#main function
def run():
    parser = argparse.ArgumentParser(description='exports speaker notes from pptx files by parsing the XML')
    parser.add_argument('-v', action='version', version='%(prog)s 1.0')
    parser.add_argument('-p',metavar='<path/to/pptx/file>',help='path to the Powerpoint 2007+ file',action='store',type=file,dest='pptxfile')
    #add more arguments here in future if you wish to expand
    args = parser.parse_args()
    #extract the pptx file as a zip archive
    #note: only extract from pptx files that you trust. they could potentially overwrite your important files.
    ZipFile(args.pptxfile).extractall(path='/tmp/', pwd=None)
    path = '/tmp/ppt/notesSlides/'
    
    notesDict = {}
    #open up the file that you wish to write to
    writepath = os.path.dirname(args.pptxfile.name)+'/'+os.path.basename(args.pptxfile.name).rsplit('.', 1)[0]+'_presenter_notes.txt'
    print writepath
    f = open(writepath, 'w')
    
    for infile in glob.glob(os.path.join(path, '*.xml')):
        #parse each XML notes file from the notes folder.
        dom = parse(infile)
        noteslist = dom.getElementsByTagName('a:t')
        #separate last element of noteslist for use as the slide marking.
        slideNumber = noteslist.pop()
        slideNumber = slideNumber.toxml().replace('<a:t>','').replace('</a:t>','')
        #start with this empty string to build the presenter note itself
        tempstring = ''
        
        for node in noteslist:
            xmlTag = node.toxml()
            xmlData=xmlTag.replace('<a:t>','').replace('</a:t>','')
            #concatenate the xmlData to the tempstring for the particular slideNumber index.
            tempstring = tempstring + xmlData
            
        #store the tempstring in the dictionary under the slide number
        notesDict[slideNumber] = tempstring
        
    #print/write the dictionary to file in sorted order by key value.
    for x in range(1,len(notesDict)):
        #print 'Slide '+str(x)+'\n'
        f.write('Slide '+str(x)+'\n')
        stringybean = notesDict[str(x)]
        f.write(stringybean.encode('ascii','ignore')+'\n')

    f.close()
    print 'file successfully written to'+'\''+writepath+'\''

if __name__ == "__main__":
    try:
        run()
    except (KeyboardInterrupt, SystemExit):
        raise
    except:
        print'something wrong happened'
        # report error and proceed


