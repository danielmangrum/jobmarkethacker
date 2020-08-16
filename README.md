# Job Market Hacker


Daniel Mangrum



This package is designed to make it easier to apply for jobs during the Economics Job Market. With this code, you can create multiple templates for your cover letters, research statements, teaching statements, and diversity statements and inject wild cards into each letter or statement to customize them to a specific job. This code will take the LaTeX files and create a separate application folder for each job application with the information you choose. 

The package contains six files (all of these files should remain in the root directory):

- JMHacker.py
- Jobs.xlsx
- Cover.tex
- Research.tex
- Teaching.tex
- Diversity.tex

Jobs.xlsx is an Excel spreadsheet where you will collect all of your job information.
Cover.tex is an example cover letter tex file for your job applications.
Research.tex is an example research statement tex file.
Teaching.tex is an example teaching statement tex file.
Diversity.tex is an example diversity statement tex file.
JMHacker.py is the python file that will compile all tex files, read the information in the Jobs spreadsheet, and create the application folders.


### Dependencies

In order to use this code, you will need to install some dependencies.

1. Python
	I use Python 3.6 but you may be able to use other versions successfully. The following packages should be installed through pip (e.g 'pip install PyPDF2' at the command line or within Anaconda or Jupyter)
   - `os`
   - `glob`
   - `pandas`
   - `numpy`
   - `PyPDF2`
   - `shutil`

1. MikTeX (or other LaTeX compiler—see http://sachaepskamp.com/wp-content/uploads/2011/10/Install.pdf)
   `pdflatex` should be accessible via your PATH.
   Packages used are (some of these aren't required but are in the example files):
    - `geometry`
    - `setspace`
    - `placeins`
    - `babel`
    - `inputenc`
    - `fancyhdr`
    - `hyperref`
    - `import`
    - `datatool`
    - `ifthen`
    - `blindtext`

1. Microsoft Excel


### Getting Started

1. To begin, open and explore the tex documents to familiarize yourself with the structure. The tex file will read information from your job spreadsheet and loop over all of the jobs you wish to compile to create a unique PDF file for each application. For example, in the Cover.tex there are various wild card variables that will be placed into the corresponding section of the cover letter. Additionally, you can create alternate versions of a section of a statement/letter (e.g. econ_academic, government, postdoc, etc). Also notice there are 'Bonus' wild cards that you can inject a unique sentence or paragraph into the statement/letter.

2. Open the Jobs Spreadsheet and explore the columns. These columns correspond to the wild cards in the tex files. You can begin to add new rows with new jobs but it is best to compile with only the sample jobs so that you know the code works. It is also recommended you keep the sample jobs in the spreadsheet so that the python file always has a job to work on. The Excel file is pre-populated with 29 columns that are used as inputs in the various tex files:

2.) Open the Jobs Spreadsheet and explore the columns. These columns correspond to the wild cards in the tex files. You can begin to add new rows with new jobs but it is best to compile with only the sample jobs so that you know the code works. It is also recommended you keep the sample jobs in the spreadsheet so that the python file always has a job to work on. The Excel file is pre-populated with 29 columns that are used as inputs in the various tex files:

Name : The name of the department and university for the cover letter (Use \\ for line breaks)

Address : Street address for cover letter (Use \\ for line breaks)

City : City for cover letter

State : State for cover letter

Zip : Zip code for cover letter 

Country : Country for cover letter (optional)

Submitted : Leave blank in order to build the application folder. Mark with 'Yes' once this application is submitted and the folder will be move to the Submitted folder

Complete : Mark with 'Yes' when the job details are completed so that the python code will process this job. Leave blank otherwise.

Deadline : Enter the date of the deadline in MM/DD/YYYY format 

Rolling : Mark with 'Yes' if this job has a rolling deadline (will be sorted to the top of Applications folder)

CoverLetterCat : Denotes which version of the cover letter to use for this application (More instructions in Cover.tex)

Position : Title of the job position (will be added verbatim in cover letter)

Contents : list of application materials that will be included (will be added verbatim in cover letter - see Cover.tex for context)

Letters : provide a list of letter writers (will be added verbatim in cover letter)

LettersFormat : Prefix to letter writers (will be added verbatim in cover letter - see Cover.tex for context)

CoverBonusGeneral : General statement to add to cover letter (will be added verbatim in cover letter - see Cover.tex for context)

CoverBonusResearch : General statement to add to the research section of cover letter (will be added verbatim in cover letter - see Cover.tex for context)

CoverBonusTeaching : General statement to add to the teaching section of cover letter (will be added verbatim in cover letter - see Cover.tex for context)

Posted : Where the job postin g was found (will be added verbatim in cover letter - see Cover.tex for context)

Interview : Availability for interviewing (will be added verbatim in cover letter - see Cover.tex for context)

Addressee : Who the letter will be addressed to (will be added verbatim in cover letter)

TeachingStatementCat : Categorize the teaching statement for Teaching.tex (leave blank to skip document)

TeachingBonus : Additional space for sentence/paragraph at the end of teaching statement in Teaching.tex

ResearchStatementCat : Categorize the research statement for Research.tex (leave blank to skip document)

ResearchBonus : Additional space for sentence/paragraph at the end of research statement in Research.tex

DiversityStatementCat : Categorize the diversity statement for Diversity.tex (leave blank to skip document)

DiversityBonus : Additional space for sentence/paragraph at the end of diversity statement in Diversity.tex

SecondPaper : Filename for additional paper/transcript/etc to be added to application folder (store in {root}/Misc folder)

ThirdPaper : Filename for additional paper/transcript/etc to be added to application folder (store in {root}/Misc folder)

You can add any new wild card to the program by adding a column to the jobs spreadsheet and then invoking the wild card in the appropriate tex file.

3. Customize the JMHacker.py file:
	Adjust the root directory to the location of the source files
	Adjust the last name variable
	Set the file location for the CV pdf
	Set the file location for the job market paper pdf

4. Run the JMHacker.py file to generate the sample application folders and the Job_export.csv. You must have all application folders closed before running the py file. Otherwise you will get an error when trying to move/delete folders.

5. Run JMHacker.py as needed to generate new application files, move newly submitted applications to the Submitted folder, and update any documents. Note that if you change the name or deadline of an application, the old folder will remain and a new folder will be created. You will need to manually delete any folder that doesn't have a matching name/deadline combination in the spreadsheet.


