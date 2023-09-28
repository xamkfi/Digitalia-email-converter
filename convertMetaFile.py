'''
Created on Sep 14, 2022

@author: digitalia-aj
'''
"""
Lets convert the metadata produced by pffexport into format understood by the pdfmarks
"""
import logwriter
import os
import re
import subprocess

def convertMetaData(root, handleMessages):
    #global beginfilenamestriped
    #logwriter.printmessage("In metadata converter, {}".format(root))
    #original_walk_dir = os.path.dirname(os.path.abspath(__file__))
    #global original_walk_dir
    try:
        """Before convert, lets combine the original metadatafiles
        More general approach"""        
        obs = os.scandir(root)
        filenames = []
        for oneobj in obs:            
            if oneobj.is_file():
                if oneobj.name.endswith(".txt") and "Message" not in oneobj.name:
                    if os.path.getsize(os.path.join(root,oneobj.name))>0:
                        filenames.append(os.path.join(root,oneobj.name))
                    #else:                    
                        #logwriter.printmessage("Excluding 0 size file {}".format(os.path.join(root,oneobj.name)))
        
        chfile = os.path.join(root, "CombinedHeaders.txt")
        modchfile = os.path.join(root, "ModifiedCombinedHeaders.txt")
        #logwriter.printmessage("txt files {}".format(filenames)) 
        """Old approach"""        
        """
        internetHeadersExists = False
        #logwriter.printmessage("1")
        if os.path.isfile(os.path.join(root,"InternetHeaders.txt")):
            filenames.append(os.path.join(root,"InternetHeaders.txt"))
            internetHeadersExists = True
        
        #logwriter.printmessage("2")
        if os.path.isfile(os.path.join(root,"OutlookHeaders.txt")):
            filenames.append(os.path.join(root,"OutlookHeaders.txt"))        
        
        if os.path.isfile(os.path.join(root,"Appointment.txt")):
            filenames.append(os.path.join(root,"Appointment.txt"))    
        
        if os.path.isfile(os.path.join(root,"Recipients.txt")):
            filenames.append(os.path.join(root,"Recipients.txt"))
            internetHeadersExists = True
        
        rfile = os.path.join(root, "Recipients.txt")        
        chfile = os.path.join(root, "CombinedHeaders.txt")
        modchfile = os.path.join(root, "ModifiedCombinedHeaders.txt")        
                
        if internetHeadersExists == False: #We only need to add recipients.txt if internetheaders does not exist
            if (os.path.isfile(rfile)):
                #logwriter.printmessage("Just another test")
                
                #But the format is wrong, below method modifies the content to To: / CC format before append.
                logwriter.printmessage("Modify {}".format(rfile))
                modifyRecipients(rfile)
                filenames.append(rfile)
        #logwriter.printmessage("Headerfiles in {} are {}".format(root, filenames))
        """
        with open(chfile, 'w') as outfile:
            if handleMessages:                
                for fname in filenames:
                    if "Meeting" not in fname:
                        with open(fname) as infile:
                            outfile.write(infile.read())
            else:
                for fname in filenames:
                    with open(fname) as infile:
                        try:
                            lines = infile.readlines()
                            i=0
                            for line in lines:
                                if "Body:" in line:                            
                                    endIndex = len(lines)
                                    #logwriter.printmessage("Lines {}".format(endIndex))
                                    startIndex = i
                                    #logwriter.printmessage("Body in line {}".format(startIndex))
                                    
                                    #logwriter.printmessage("Body in metadata {} delete from {} to {}".format(chfile, startIndex, endIndex))
                                    del lines[startIndex:endIndex]
                                    #logwriter.printmessage("Deletion OK")
                                    break
                                i+=1
                            for line in lines:                        
                                outfile.write(line)
                            outfile.close()
                        #logwriter.printmessage("Success")
                        except:
                            pass
                        
                    
        if os.path.isfile(chfile):
            #logwriter.printmessage("Combined file found - convert")
            #convertMetaFile(chfile, modchfile)
            #logwriter.printmessage("Original walk dir is now {}".format(original_walk_dir))
            jarpath = os.path.join(os.path.dirname(os.path.abspath(__file__)), "convertmetadata.jar")
            #logwriter.printmessage ("JAR: {} {} {}".format(jarpath, temptail, root))
            #subprocess.call(["java", "-jar", jarpath, chfile, modchfile, csvfile, temptail])
            #javacmd = ["java", "-jar", jarpath, chfile, modchfile, csvfile, temptail]
            javacmd = ["java", "-jar", jarpath, chfile, modchfile]
            cmdProcess = subprocess.Popen(javacmd, stdout=subprocess.PIPE)
            result = cmdProcess.communicate()[0] 
            
            #logwriter.printmessage("Java process results {}".format(result))
            #isoheader = str(combheaderout.decode('iso-8859-1'))
            #logwriter.printmessage("Testi2"+isoheader)


        """Directory"+sep+"From"+sep+"To"+sep+"CC"+sep+"Subject"+sep+"Date\n"""

        #headerlist = []

    except OSError:
        pass


    #Reads the final header list and formulates a string to be written into csv file
    #finalstring = ';'.join(headerlist)
    #logwriter.printmessage(finalstring)
    #writescv(finalstring)
    #logwriter.printmessage ("FINAL HEADER list == "+str(headerlist))
    #logwriter.printmessage ("Metadata conversion and csv formulation ended in %s seconds" %(time.time()-starttime))
    return

"""
This is currently not used but can later on
Replace convertmetadata.jar if datetime conversions
are added to this
"""
def convertMetaFile(combinedFile, modedFile):
    #logwriter.printmessage("Moded file = {}".format(modedFile))
    contentLines = []
    inputFile = open(combinedFile, "r")
    allLines = inputFile.readlines()
    #counter = 1
    templine = ""
    for line in allLines: #Does not include the last line currently        
        if re.match(r'^\S', line):
            if templine!="":
                contentLines.append(templine)
                templine = ""
            #Metadata line starts with this kind of notation
            templine=line
        elif re.match(r'^\s+', line) and templine !="": #Starts with white space(s) add to templine                
            templine+=line
    if templine !="": #Adds the last filled line
        contentLines.append(templine)    
    #logwriter.printmessage("Done seeking")
    inputFile.close()
    
    #allLines.clear()
    #output = open(modedFile, "x")    
    #output.write("Test")
    #output.close()   
    output = open(modedFile, 'w')
    #logwriter.printmessage("Content line count = {}".format(len(contentLines)))
    for oneline in contentLines:
        oneline = oneline.replace("\n", " ").replace("\t", "").strip()
        head, value = oneline.split(":", 1)
        head = "/"+head
        value = "("+value.strip()+")"
        #output.write("Test")
        output.write("{} {}\n".format(head, value))
        #print("{} {}".format(head, value))
        #counter+=1

    output.close()
    
    return

"""
Converts recipient.txt file if needed
"""
def modifyRecipients(rfile):
    logwriter.printmessage("MODIFY: {}".format(rfile))
    toRecipients = []
    ccRecipients = []
    tempLines = []
    with open(rfile, "r") as inputFile:
        counter = 1;
        for line in inputFile:
            if line!="\n":
                tempLines.append(line)
            else:
                #print(str(counter)+":"+str(tempLines))
                appendName = ""
                for oneline in tempLines:
                    oneline = oneline.replace("\t", "").strip()
                    head, value = oneline.split(":")
                    
                    #logwriter.printmessage("{},{}".format(head, value))
                    if "display" in head.lower():                    
                        appendName = value.replace("\t", "").replace("\n", "")
                    if "To" in value:
                        #print(appendName)
                        if appendName != "":
                            toRecipients.append(appendName)
                        appendName = ""
                    if "CC" in value:
                        if appendName != "":
                            ccRecipients.append(appendName)
                        appendName = ""                                    
                counter+=1
                tempLines=[]
    toRecipients = str(toRecipients).replace("[","").replace("]","")
    ccRecipients = str(ccRecipients).replace("[","").replace("]","")
    logwriter.printmessage(toRecipients)
    logwriter.printmessage(ccRecipients)
    input.close()
    
    output = open(rfile, "w")
    if len(toRecipients)>0:
        output.write("To:"+str(toRecipients)+"\n")
    if len(ccRecipients)>0:
        output.write("Cc:"+str(ccRecipients)+"\n")
    output.write("\n")
    output.close      
    
    return

