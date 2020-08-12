import os
import glob
import pandas as pd 
import numpy as np
from PyPDF2 import PdfFileReader, PdfFileWriter
from shutil import copyfile
import shutil



# Set the root directory for the Job Market Hacker python file and tex files
root_dir = 'C:/Users/Daniel/Dropbox/GitHub/jobmarkethacker/'

#Set the name you want as a prefix to your documents in the application folders
last_name = 'Mangrum'


# Set the location of your CV and Job Market Paper. These will be copied into 
# each application folder
cv_file = '{}/Misc/MangrumCV.pdf'.format(root_dir)
jmp_file = '{}/Misc/Mangrum_JMP.pdf'.format(root_dir)


# Do you want to delete and recreate application folders that already exist?
# The code will copy application folders to a 'Submitted' folder after you
# submit them.
# NOTE: If an application is marked anything in the Submitted row, this code 
# will skip that entry in the Jobs spreadsheet
delete_existing_folders = 1


# Compile LaTeX documents?
# This is recommended as to create the appropriately sized repeated PDF files
compile = 1



### Begin processing code ###


# Create a few directories
try:
    os.mkdir('{}/Applications'.format(root_dir))
except:
    pass
try:
    os.mkdir('{}/Applications/Submitted'.format(root_dir))
except:
    pass

# Create CSV of Job details from the main Jobs.xlsx
Jobs_xlsx = pd.read_excel('{}/Jobs.xlsx'.format(root_dir), sheet_name='Cover Letter Information', converters={'Zip':str})
Jobs_xlsx = Jobs_xlsx.replace('nan', '')
Jobs_xlsx = Jobs_xlsx[~Jobs_xlsx.Name.isna()]
Jobs_xlsx.Submitted = Jobs_xlsx.Submitted.astype(str)
Jobs_xlsx.Submitted = Jobs_xlsx.Submitted.replace("nan","")
Jobs_list = Jobs_xlsx.copy()
Jobs_xlsx = Jobs_xlsx[~Jobs_xlsx.Complete.isna()]




# Move newly submitted applications into the Submitted folder

Jobs_xlsx.reset_index(inplace=True)
jobs_n = len(Jobs_xlsx)

# Loop over the list of completed job profiles
for job in range(0,jobs_n):
    # Select all jobs markted Yes in Submitted column
    if Jobs_xlsx.Submitted[job] == 'Yes':
        # Format the name of the folder and move the folder from Applications
        # to Submitted
        try:
            due = Jobs_xlsx.Deadline[job].strftime("%Y-%m-%d")
            if Jobs_xlsx.Rolling[job] == 'Yes':
                due = '0Rolling'
            name = Jobs_xlsx.Name[job]
            name = name.replace(" \\\\ ", "] [")
            name = name.replace("\&", "&")
            name_save = name.replace(" ", "")
            old_path = '{}/Applications/[{}] [{}]'.format(root_dir,due,name)
            new_path = '{}/Applications/Submitted/[{}] [{}]'.format(root_dir,due,name)
            shutil.move(old_path,new_path)
            print('Moved [{}_{}] to Submitted!'.format(due,name))
        except:
            print('Unable to move {}_{} to Submitted!'.format(due,name))
            continue


# Keep only jobs that have not been submitted (but are Complete)
Jobs_xlsx = Jobs_xlsx[Jobs_xlsx.Submitted != 'Yes']

# Save job details to a csv to be read by pdflatex
Jobs_xlsx.to_csv(r'{}/Jobs_export.csv'.format(root_dir),index=False, float_format="%.0f")





print('')
print('')

# Read Job details from the Jobs spreadsheet
Job_details = pd.read_csv('{}/Jobs_export.csv'.format(root_dir))

jobs_n = len(Job_details) 

if compile == 1:
    # Compile Documents in LaTeX
    os.system('pdflatex  Cover.tex' )
    os.system('pdflatex  Research.tex' )
    os.system('pdflatex  Teaching.tex' )
    os.system('pdflatex  Diversity.tex' )

#Remove any log, aux, or out files keeping only pdf files
    fileList = glob.glob('{}/*.log'.format(root_dir))
    for item in fileList:
        try:
            os.remove(item)
        except:
            pass
    fileList = glob.glob('{}/*.aux'.format(root_dir))
    for item in fileList:
        try:
            os.remove(item)
        except:
            pass
    fileList = glob.glob('{}/*.out'.format(root_dir))
    for item in fileList:
        try:
            os.remove(item)
        except:
            pass



# Path for batch pdf of application documents
RS_path = '{}/Research.pdf'.format(root_dir)
TS_path = '{}/Teaching.pdf'.format(root_dir)
DS_path = '{}/Diversity.pdf'.format(root_dir)
CL_path = '{}/Cover.pdf'.format(root_dir)


# Read in each PDF
CL_pdf = PdfFileReader(CL_path)
TS_pdf = PdfFileReader(TS_path)
RS_pdf = PdfFileReader(RS_path)
DS_pdf = PdfFileReader(DS_path)

# The next section of code will find the end of each sub-document within the 
# batch PDF so that the documents can be spliced.
# NOTE: Cover letters are designed to only be one page. If you have cover letters
# longer than one page, this will fail. You can repeat the same logic used for 
# the other documents if you need.

# NOTE: This syntax will fail for any document that is only 1 page, hence the try/except syntax.
try:
    TS_page_end = TS_pdf.trailer["/Root"]["/PageLabels"]["/Nums"]
    while {'/S': '/D'} in TS_page_end: TS_page_end.remove({'/S': '/D'})
    while 0 in TS_page_end: TS_page_end.remove(0)
except:
    pass
try:
    RS_page_end = RS_pdf.trailer["/Root"]["/PageLabels"]["/Nums"]
    while {'/S': '/D'} in RS_page_end: RS_page_end.remove({'/S': '/D'})
    while 0 in RS_page_end: RS_page_end.remove(0)
except:
    pass
try:
    DS_page_end = DS_pdf.trailer["/Root"]["/PageLabels"]["/Nums"]
    while {'/S': '/D'} in DS_page_end: DS_page_end.remove({'/S': '/D'})
    while 0 in DS_page_end: DS_page_end.remove(0)
except:
    pass





TS_start = 0
RS_start = 0
DS_start = 0

TS_count = 0
RS_count = 0
DS_count = 0


### Begin the loop over each job application ###

# Loop over each entry in the job details CSV
for page in range(0,jobs_n):
     #Store the name of the job from header "Name"  from Job Details CSV
    name = Job_details.Name[page]
    print('')
    # Process the name string for creating the application folder
    name = name.replace("\&", "&")
    name = name.replace(" \\\\ ", "] [")
    print('[{}]'.format(name))
    name_save = name.replace(" ", "")
    #Store the "Deadline" header from Job Details CSV (due date of application)
    due = Job_details.Deadline[page]
    #Remove date slashes and replace with underscores
    due = due.replace("/", "_")
    
    # The naming convention allows for sort by due date. All applications 
    # with a rolling deadline begin with 0 so as to be placed at the top
    if Job_details.Rolling[page] == 'Yes':
        due = '0Rolling'
    
    # Where to save each application. The naming format puts due date first 
    # (for sort purposes), then Name from Job Details
    savepath = '{}/Applications/[{}] [{}]'.format(root_dir,due,name)
    # PDF filename to save the split cover letters
    output_filename = '{}/{}_Cover_Letter.pdf'.format(savepath,last_name)

    # Create directory for Application
    if os.path.exists(savepath):
        if delete_existing_folders == 1:
            print('   Previous application for [{}] exists. Deleting!'.format(name))
            shutil.rmtree(savepath)
            os.mkdir(savepath)
        else:
            print('   Application for [{}] already exists'.format(name))
            continue   
            # If application folder exists, continue
    else:
        os.mkdir(savepath)
       

   
    
    # Begin by splitting cover letters into individual pages
    pdf_writer_CL = PdfFileWriter()
    pdf_writer_CL.addPage(CL_pdf.getPage(page))
    
    
    
    #Split Cover Letter pdf into pages and write the page
    #ONLY WORKS WITH SINGLE PAGE Cover Letter
    with open(output_filename, 'wb') as out:
        pdf_writer_CL.write(out)


    ### Begin processing statement documents. In order for the loop to process
    # a document, there must be something in the XXXXStatement column of the
    # Job_details csv
    
    # Teaching Statement
    if ~Job_details.TeachingStatementCat.isna()[page]:   
        try:
            # Load the page number for the end of subdocument
            TS_end = TS_page_end[TS_count]
        except:
            # If single document, use the number of pages of the main document
            TS_end = TS_pdf.getNumPages() 
        # Create blank pdf    
        pdf_writer_TS = PdfFileWriter()
        # Append each page to the blank PDF until it reachest the end
        while TS_start < TS_end:
            pdf_writer_TS.addPage(TS_pdf.getPage(TS_start))
            TS_start = TS_start + 1
        
        # Output the PDF into the current application folder
        TS_output_filename = '{}/{}_Teaching.pdf'.format(savepath,last_name)
        with open(TS_output_filename, 'wb') as out:
            pdf_writer_TS.write(out)
        # Advance the end page counter for the next application
        TS_start = TS_end
        TS_count = TS_count + 1

    # Research Statement
    if ~Job_details.ResearchStatementCat.isna()[page]:
        try:
            # Load the page number for the end of subdocument
            RS_end = RS_page_end[RS_count]
        except:
            # If single document, use the number of pages of the main document
            RS_end = RS_pdf.getNumPages() 
        # Create blank pdf    
        pdf_writer_RS = PdfFileWriter()
        # Append each page to the blank PDF until it reachest the end
        while RS_start < RS_end:
            pdf_writer_RS.addPage(RS_pdf.getPage(RS_start))
            RS_start = RS_start + 1
        
        # Output the PDF into the current application folder        
        RS_output_filename = '{}/{}_Research.pdf'.format(savepath,last_name)
        with open(RS_output_filename, 'wb') as out:
            pdf_writer_RS.write(out)
        # Advance the end page counter for the next application        
        RS_start = RS_end
        RS_count = RS_count + 1

    # Diversity Statement
    if ~Job_details.DiversityStatementCat.isna()[page]:
        try:
            # Load the page number for the end of subdocument            
            DS_end = DS_page_end[DS_count]
        except:
            # If single document, use the number of pages of the main document            
            DS_end = DS_pdf.getNumPages() 
        # Create blank pdf     
        pdf_writer_DS = PdfFileWriter()
        # Append each page to the blank PDF until it reachest the end
        while DS_start < DS_end:
            pdf_writer_DS.addPage(DS_pdf.getPage(DS_start))
            DS_start = DS_start + 1
        # Output the PDF into the current application folder        
        DS_output_filename = '{}/{}_Diversity.pdf'.format(savepath,last_name)
        with open(DS_output_filename, 'wb') as out:
            pdf_writer_DS.write(out)
        # Advance the end page counter for the next application        
        DS_start = DS_end
        DS_count = DS_count + 1




    # Copy CV into Application folder
    dst = '{}/{}_CV.pdf'.format(savepath,last_name)
    copyfile(cv_file,dst)


    # Copy Job Market Paper  
    dst = '{}/Mangrum_JMP.pdf'.format(savepath,name_save)
    copyfile(jmp_file,dst)

    # I've added the ability to copy additional papers/documents to the 
    # application folder using the SecondPaper and ThirdPaper columns
    # in the Job_details csv. 
    try:
        if (Job_details.SecondPaper[page]!= ""):
            paper_file = Job_details.SecondPaper[page]
            src = '{}/Misc/{}'.format(root_dir,paper_file)
            dst = '{}/{}'.format(savepath,paper_file)
            copyfile(src,dst)
    except:
         if not np.isnan(Job_details.SecondPaper[page]):
             paper_file = Job_details.SecondPaper[page]
             src = '{}/Misc/{}'.format(paper_file)
             dst = '{}/{}'.format(savepath,paper_file)
             copyfile(src,dst) 
    try:
         if (Job_details.ThirdPaper[page]!= ""):
            paper_file = Job_details.ThirdPaper[page]
            src = '{}/Misc/{}'.format(paper_file)
            dst = '{}/{}'.format(savepath,paper_file)
            copyfile(src,dst)
    except:
         if not np.isnan(Job_details.ThirdPaper[page]):
             paper_file = Job_details.ThirdPaper[page]
             src = '{}/Misc/{}'.format(paper_file)
             dst = '{}/{}'.format(savepath,paper_file)
             copyfile(src,dst) 

    print('   [{}] Application created!'.format(name))

# Remove batch pdf files
fileList = glob.glob('{}/*.pdf'.format(root_dir))
for item in fileList:
    try:
        os.remove(item)
    except:
        pass
        
# Remove batch aux files
fileList = glob.glob('{}/*.aux'.format(root_dir))
for item in fileList:
    try:
        os.remove(item)
    except:
        pass
        