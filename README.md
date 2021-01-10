# BaseTextScraperv.0.6
The idea is to take a table out of a messy text file that holds the data. 
Unfortunately, there is a lot of whitespace and several structures such as multiples of =, _ which need to be cleaned up. 
At the moment the code is a lot of copy and paste of one way to extract the data properly sorted into int and str in an excel table. 
For now the code scrapes the data for the 6 working days of the week from the file where the text files are stored but this is only a temporary solution. 
Next versions should include: 
- several loops that sort the txt file into those that include the category 109h and those that do not 
- several loops that do what is now done through copy and pasting similar code that is then fed new locations for where to take the data from into loops that do so without specific direction to the exact location of the data in the file, but rather by assigning the index through logic
- the scraper should read all the files in the folder and put the data of however many files are in the folder into the created excel table

The versions after that should include: 
- visualization of the most relevant data being scraped 
- a second worksheet with a calculation that returns the true number of daily orders arriving in the magazine of 109a, taking account of orders put in outside of working hours
