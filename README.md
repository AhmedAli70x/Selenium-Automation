# Selenium-Automation

The program has config file where user can add the week range and edit the running mode (default or failed urls), the main file has the code functionality and data processing.
This script get the patients data from geodes website and export it to excel sheet.
The program check for required packeg first, download a package if it is not installed.
For each url, the program scrape the number of patients, them them to local excell sheet.
Error handling is used to handle failed to scrap urls, then these urls are saved to failed.txt file. They can be scraped later separately using the config file.
### Running CMD
<img width="966" alt="image" src="https://user-images.githubusercontent.com/71847656/160706012-4ca457cc-4691-44c3-b7b7-07e154ea4e13.png">


### CSV result
<img width="1118" alt="image" src="https://user-images.githubusercontent.com/71847656/160706171-a91fe1d4-fb71-4dce-be08-de9d0509aceb.png">
