#Goals:
#Create new name for zip file based on name of Assembly guide.
#Create name based on CDR file if no guide exists
#Create thumbnail based on first page of Assmbly guide
#Create thumbnail based on ai or svg if no assembly guide
#make table for dokuwiki and csv of:
    #old filename, new filename, path, date of thumbnail file
#optional: Delete thumbnail.db

# importing required modules
#Working with Zip Files
from zipfile import ZipFile
# For renaming
import os

#for exiting if set files found
import sys

#time manipulation for getting modified times in zip file to track newer projects
import datetime

#for pdf image manipulation
#sudo apt install python3-pip
#pip3 install Wand
from wand.image import Image
from wand.color import Color

#Misc fixes:
#wand.exceptions.PolicyError: not authorized
#solved by changing:
#<policy domain="coder" rights="none" pattern="PDF" />
#to
#<policy domain="coder" rights="read" pattern="PDF" />
#in
#/etc/ImageMagick-6/policy.xml

# path to folder which needs to be read
directory = '/media/sf_SharedVirtualBox'
#directory = '/media/sf_SharedVirtualBox/Intarsia'
#directory = '/media/xenth/MIDNIGHT'
#directory = '/media/xenth/MIDNIGHT/Intarsia'

def get_all_file_paths(directory):

    # initializing empty file paths list
    file_paths = []

    # crawling through directory and subdirectories
    for root, directories, files in os.walk(directory):
        for filename in files:
            # join the two strings in order to form the full filepath.
            filepath = os.path.join(root, filename)
            file_paths.append(filepath)

    # returning all file paths
    return file_paths

def convert_pdf(filename, output_path, resolution=72, pages_captured=0):
    """ Convert a PDF into images.

        All the pages will give a single png file with format:
        {pdf_filename}-{page_number}.png

        The function removes the alpha channel from the image and
        replace it with a white background.
    """

    all_pages = Image(filename=filename, resolution=resolution)
    for i, page in enumerate(all_pages.sequence):
        with Image(page) as img:
            img.format = 'png'
            img.background_color = Color('white')
            img.alpha_channel = 'remove'

            #same as starting file minus extension
            image_filename = os.path.splitext(os.path.basename(filename))[0]
            image_filename = image_filename.lower().replace(' assembly guide', '').replace(' assembly', '').replace(' ', '_')
            #neat for page number but not needed in this case
            if (pages_captured>0):
            	  image_filename = '{}-{}.png'.format(image_filename.title(),i)
            else:
            	  image_filename = '{}.png'.format(image_filename.title())
            image_filename = os.path.join(output_path, image_filename)

            img.save(filename=image_filename)
            if i == pages_captured:
                all_pages.destroy()
                return image_filename

def convert_image(filename, output_path, resolution=72, pages_captured=0):
    """ Convert a PDF into images.

        All the pages will give a single png file with format:
        {pdf_filename}-{page_number}.png

        The function removes the alpha channel from the image and
        replace it with a white background.
    """

    all_pages = Image(filename=filename, resolution=resolution)
    for i, page in enumerate(all_pages.sequence):
        with Image(page) as img:
            img.format = 'png'
            img.background_color = Color('white')
            img.alpha_channel = 'remove'

            #same as starting file minus extension
            image_filename = os.path.splitext(os.path.basename(filename))[0]
            image_filename = image_filename.lower().replace('_svg', '').replace(' ', '_')
            #neat for page number but not needed in this case
            if (pages_captured>0):
            	  image_filename = '{}-{}.png'.format(image_filename.title(),i)
            else:
            	  image_filename = '{}.png'.format(image_filename.title())
            image_filename = os.path.join(output_path, image_filename)

            img.save(filename=image_filename)
            if i == pages_captured:
                all_pages.destroy()
                return image_filename

def archive_check():
    #Reads recursively through all the directories in "directory" and scans for zip files inside zip files
    # calling function to get all file paths in the directory
    file_paths = get_all_file_paths(directory)

    #No zip sets found
    zip_set = False
    info=''

    # printing the list of all zipped files
    for file_name in file_paths:
        if file_name.lower().endswith('.zip'):
            #print(file_name)
            # opening the zip file in READ mode
            with ZipFile(file_name, 'r') as zip:
                # printing all the contents of the zip file
                #zip.printdir()
                name_list=''
                #Only log if set file
                log = False

                for zfile_name in zip.namelist():
                    name_list+=zfile_name + "\n"
                    if zfile_name.lower().endswith('.zip'):
                        zip_set = True
                        log = True
                if log == True:
                        info+="\n*** " + file_name + " *** is an archive set.  Manual intervention required.\n" + name_list
    if zip_set:
        #setup logging files for problems
        man = open("!!!Zipped_files.txt","w")
        man.write(info)
        man.close()
        ignore_zip = input("Archive Files found.  Read !!!Zipped_files.txt for more information.\nContinue processing? Y/[N]: ")  # Python 3
        if ignore_zip != "y" and ignore_zip != "Y":
            sys.exit("Processing stopped")
    else:
        print("No zip sets found.  Continuing with cleanup")


def main():
    #rename dictionary
    rename = dict()

    #setup logging files for problems
    oopsies = open("undo.sh","w")

    #setup logging files for problems
    man = open("!!!manual_intervention.txt","w")

    #setup DokuWiki full table
    doku = open("makeCNC_dokuwiki.txt","w")

    #setup CSV full table
    csv = open("makeCNC.csv","w")

    #Headers  Thumbnail, Title, File, Original MakeCNC Filename, Assembly Instructions, Modified Date, Search (in title of wiki)
    csv.write("Thumbnail,Title,File,Original MakeCNC Filename,Assembly Instructions,Modified Date,Search" + "\n")
    doku.write("<sortable 2>\n")
    doku.write("^Thumbnail^Title^File^Original MakeCNC Filename^Assembly Instructions^Modified Date^\n")

    # calling function to get all file paths in the directory
    file_paths = get_all_file_paths(directory)

    # printing the list of all zipped files
    for file_name in file_paths:
        if file_name.lower().endswith('.zip'):
            print("Processing: " + file_name)
            # opening the zip file in READ mode
            with ZipFile(file_name, 'r') as zip:
                # printing all the contents of the zip file
                #zip.printdir()
                check = False
                cdr = ''
                for zfile_name in zip.namelist():
                	  #show info about zip file
                    info = zip.infolist()[0]
                    #Modified:
                    when = str(datetime.datetime(*info.date_time))
                    #System (0 = Windows, 3 = Unix:
                    how = str(info.create_system)

                    if zfile_name.lower().endswith('.pdf'):
                        #Assembly Instructions found.  Probably has a good name minus " Assembly Guide.pdf"
                        if "assembly" in zfile_name.lower():

                            #print(os.path.abspath(file_name))   /media/sf_SharedVirtualBox/Waterkelpie.zip
                            #print(os.path.basename(file_name))  Waterkelpie.zip
                            #print(os.path.dirname(file_name))   /media/sf_SharedVirtualBox

                            #extract the PDF to a temporary location
                            zip.extract(zfile_name, "/tmp/")
                            #convert to a png
                            png_file = convert_pdf("/tmp/"+zfile_name, os.path.dirname(file_name))
                            os.remove("/tmp/"+zfile_name)

                            #cleaned up name
                            new_name = os.path.splitext(os.path.basename(zfile_name))[0]
                            new_name = new_name.lower().replace(' assembly guide', '').replace(' assembly', '').replace(' ', '_')
                            search = "https://www.makecnc.com/search.php?pg=1&posted=1&stype=&stext=" + new_name.replace('_', '+') + "&bsubmit=Search"
                            title_name = new_name.replace('_', ' ').title()
                            new_name = '{}.zip'.format(new_name.title())

                            #Thumbnail, Title, File, Original MakeCNC Filename, Assembly Instructions, Modified Date, Search (in title of wiki)
                            csv.write(os.path.basename(png_file))          #Thumbnail          
                            csv.write("," + title_name)                    #Title
                            csv.write(","+ new_name)                       #Renamed Filename
                            csv.write("," +  os.path.basename(file_name))  #Original Filename  
                            csv.write(",Yes")                              #Assembly or Intarsia
                            csv.write("," + when)                          #Modified Date      
                            csv.write("," + search)                        #Search URL
                            csv.write("\n")                                #End CSV Row

                            #Thumbnail, Title & search, File, Original MakeCNC Filename, Assembly Instructions, Modified Date
                            doku.write("| {{makecnc:" + os.path.basename(png_file) + "?50}} ") #Thumbnail
                            doku.write("| [[" + search + "|" + title_name + "]] ")             #Title and Search
                            doku.write("| {{makecnc:"+ new_name + " |}}")                      #Renamed Filename
                            doku.write("| " + os.path.basename(file_name) + " ")               #Original Filename
                            doku.write("| Yes ")                                               #Assembly or Intarsia
                            doku.write("| " + when + " ")                                      #Modified Date
                            doku.write("|\n")                                                  #End DokuWiki Row

                            #rename zip file to cleaner name
                            rename[os.path.abspath(file_name)] = os.path.dirname(file_name)+"/"+new_name
                            oopsies.write("rm " + png_file + "\n")
                            oopsies.write("mv " + os.path.dirname(file_name)+"/"+new_name + " " + os.path.abspath(file_name) + "\n")
                            check = True
                        #Intarsia found.  Manual Intervention needed
                        if "intarsia" in zfile_name.lower():
                            #update name to Intarsia_(kids stuff|critters|whatevs)_name.zip|png
                            for tfile_name in zip.namelist():
                            	  if tfile_name.lower().endswith('.svg'):
				                            #extract the PDF to a temporary location
				                            zip.extract(tfile_name, "/tmp/")
				                            #convert to a png
				                            png_file = convert_image("/tmp/"+tfile_name, os.path.dirname(file_name), 20)
				                            os.remove("/tmp/"+tfile_name)
				                            
								                    #cleaned up name
				                            new_name = os.path.splitext(os.path.basename(tfile_name))[0]
				                            new_name = new_name.lower().replace('_svg', '').replace(' ', '_')
				                            search = "https://www.makecnc.com/search.php?pg=1&posted=1&stype=&stext=" + new_name.replace('_', '+') + "&bsubmit=Search"
				                            title_name = new_name.replace('_', ' ').title()
				                            new_name = '{}.zip'.format(new_name.title())
                            #Thumbnail, Title, File, Original MakeCNC Filename, Assembly Instructions, Modified Date, Search (in title of wiki)
                            csv.write(os.path.basename(png_file))          #Thumbnail          
                            csv.write("," + title_name)                    #Title
                            csv.write(","+ new_name)                       #Renamed Filename
                            csv.write("," +  os.path.basename(file_name))  #Original Filename  
                            csv.write(",Intarsia")                         #Assembly or Intarsia
                            csv.write("," + when)                          #Modified Date      
                            csv.write("," + search)                        #Search URL
                            csv.write("\n")                                #End CSV Row

                            #Thumbnail, Title & search, File, Original MakeCNC Filename, Assembly Instructions, Modified Date
                            doku.write("| {{makecnc:" + os.path.basename(png_file) + "?50}} ") #Thumbnail
                            doku.write("| [[" + search + "|" + title_name + "]] ")             #Title and Search
                            doku.write("| {{makecnc:"+ new_name + " |}}")                      #Renamed Filename
                            doku.write("| " + os.path.basename(file_name) + " ")               #Original Filename
                            doku.write("| Intarsia ")                                          #Assembly or Intarsia
                            doku.write("| " + when + " ")                                      #Modified Date
                            doku.write("|\n")                                                  #End DokuWiki Row

                            #rename zip file to cleaner name
                            rename[os.path.abspath(file_name)] = os.path.dirname(file_name)+"/"+new_name
                            oopsies.write("rm " + png_file + "\n")
                            oopsies.write("mv " + os.path.dirname(file_name)+"/"+new_name + " " + os.path.abspath(file_name) + "\n")
				                    
                            check = True
                if check == False:
                    search = "https://www.makecnc.com/search.php?pg=1&posted=1&stype=&stext=" + os.path.basename(file_name).replace('.zip', '').replace('_', '+').replace('-', '+') + "&bsubmit=Search"
                    #Thumbnail, Title, File, Original MakeCNC Filename, Assembly Instructions, Modified Date, Search (in title of wiki)
                    csv.write("")                                  #Thumbnail          
                    csv.write(",")                                 #Title
                    csv.write(",")                                 #Renamed Filename
                    csv.write("," +  os.path.basename(file_name))  #Original Filename  
                    csv.write(",Unknown")                          #Assembly or Intarsia
                    csv.write("," + when)                          #Modified Date      
                    csv.write(",")                                 #Search URL
                    csv.write("\n")                                #End CSV Row

                    #Thumbnail, Title & search, File, Original MakeCNC Filename, Assembly Instructions, Modified Date
                    doku.write("| ")                                                   #Thumbnail
                    doku.write("| [[" + search + "|Search for Product]] ")             #Title and Search
                    doku.write("| ")                                                   #Renamed Filename
                    doku.write("| " + os.path.basename(file_name) + " ")               #Original Filename
                    doku.write("| Unknown ")                                           #Assembly or Intarsia
                    doku.write("| " + when + " ")                                      #Modified Date
                    doku.write("|\n")                                                  #End DokuWiki Row
                    print(file_name + " has no known instruction files.  Manual intervention required.\n")
                    man.write(file_name + " has no known instruction files.  Manual intervention required.\n")
    oopsies.close()
    man.close()
    doku.write("</sortable>\n")
    doku.close()
    csv.close()
    for loc in rename:
    	  #print(loc + ' now ' + rename[loc])
    	  os.rename(loc, rename[loc])

if __name__ == "__main__":
    archive_check()
    main()

