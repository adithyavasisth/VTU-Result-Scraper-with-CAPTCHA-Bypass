# VTU Result Scraper with Captcha Bypassing
A Web Scraper to scrape the results from results.vtu.ac.in and automatically fills in the Captcha using Tesseract. 

Input : 

  1.{College Code}  
  
  2.{Year of Joining(yy)}  
  
  3.{Starting USN} - {Ending USN} (USNs for the whole branch can be entered or only fixed range)
  
  4.{Semester for which the Results have been released recently}
  
Output : 

  The output is in Excel Format. 
   1. The Excel Sheet of the Marks
   2. The Excel Sheet with the Grades and GPA
   3. The Rank list for the given range of USNs.

Requirements : 

  1.Python(Set the PATH as well)
  2.Download/Install Chromedriver and Tesseract and change the path of these files inside the code 'scraper.py'.
  3.pip install {the-required-packages-which-have-been-imported}

This works for any VTU college. The code has to be updated every semester as changes are being made continuously with the Website or with the schemes are being changed.

Currently works for 7th semester. Shall be updated frequently as results release for every semester.
