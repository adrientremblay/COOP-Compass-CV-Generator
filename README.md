# COOP-Compass-CV-Generator
Uses [Selenium with Python](https://selenium-python.readthedocs.io/) to scrape Concordia University's COOP Compass job listings for hiring manager info.
This info is then filled into a word document template and exported to PDF using Powershell.

## Function
scraper.py uses chromedriver to launch a chrome window controlled by Selenium Python.  Selenium then navigates to the MyConcordia Website
and clicks around until it arrives at the job listings page on COOP Compass.  Using the parameter used to launch scraper.py the correct job
posting is opened and the hiring manager data is stored in a temporary txt file called temp.txt.  From there, fill.ps1 is launched where is reads temp.txt
and fills in the data into master_coverletter_template.docx with the use of special tag strings.  Unused tags are removed, then the new 
contents are saved to master_coverletter_output.docx then master_coverletter_template.pdf.

## Running
**This can only be run on windows due to the limitations of running Powershell scrips on other operating systems!**

- Create and activate a python venv using requirements.txt
- set envirnoment variables USERNAME_MYCONCORDIA and PASSWORD_MYCONCORDIA to your login details
  ```
  $Env:USERNAME_MYCONCORDIA = "<username>"
  $Env:PASSWORD_MYCONCORDIA = "<password>"
  ``` 
  in Powershell
- Launch scraper/scraper.py <job id> where <job id> is the ID of the listing
  - ex: `python3 scraper.py 38179`
