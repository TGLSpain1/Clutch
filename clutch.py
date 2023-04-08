#LAUNCH BROWSER MANUALLY IN COMMAND LINE WITH: chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\selenium\ChromeProfile"

# SELENIUM IMPORTS TO BE ABLE TO FIND ELEMENTS, AND SET WAIT CONDITIONS FOR SCRIPT TO CONTINUE UNTIL SPECIFIC ELEMENTS ARE VISIBLE. THEY ARE IMPORTED FROM SELENIUM BUT
# WILL BE USED BY UNDETECTED CHROMEDRIVER
from selenium.webdriver.common.by import By

# PANDAS LIBRARY TO CREATE DATAFRAMES
import pandas as pd

from functions_and_variables import set_up_driver, set_total_company_number, set_results_page_number, set_results_page_company, update_tracked_variables, random_pause, chrome_driver, base_url, email_text, get_company_blocks, check_text_field_empty, get_number_of_results_pages, create_results_pages_urls, get_company_name, get_company_website, get_company_location, get_employees_number, get_service_focus, get_contact_url, get_all_company_info, wait_for_submit, fill_company, select_partnership, write_email, deselect_email_and_shortlist, send_email, update_email_sent_excel, send_email_process, navigate_results_page, check_excel_exists

# TRACKING VARIABLE FUNCTIONS
total_company_number = set_total_company_number()
results_page_number = set_results_page_number()
results_page_company = set_results_page_company()
print(f"Current company_number is {total_company_number}")
print(f"Current results_page_number is {results_page_number}")
print(f"Current results_company_number is {results_page_company}")

# DRIVER SET UP FUNCTIONS
driver = set_up_driver()
driver.get(base_url)

# FUNCTIONS FOR GENERATING ALL RESULTS PAGES URLS
number_of_results_pages = get_number_of_results_pages(driver)
print(f"Number of results pages is {number_of_results_pages}")
results_pages_urls = create_results_pages_urls(number_of_results_pages)

# CHECKING IF EXCEL EXISTS
check_excel_exists()

# FUNCTIONS FOR GETTING ALL THE INFORMATIN FOR EACH COMPANY AND FUNCTIONS FOR FILLING OUt CONTACT FORM AND SENDING EMAIL
# for results_pages_url in results_pages_urls[results_page_number:(number_of_results_pages + 1)]:
for results_pages_url in results_pages_urls[0:10]:
	companies = []
	print(f"Working with {results_pages_url}")
	navigate_results_page(results_pages_url, driver)

	companies_li_elements = get_company_blocks(driver)
	
	for company_li_element in companies_li_elements:	

		# GET DETAILS FOR EACH COMPANY AND SAVE TO DATA_FRAME AND APPEND TO LIST OF DICTIONARIES
		company_dict = get_all_company_info(company_li_element)
		companies.append(company_dict)

	# SAVE LIST OF DICTIONARIES WITH COMPANY DETAILS TO EXCEL
	unsent_companies_dataframe = pd.DataFrame(companies)
	filename = f"clutch_companies_{results_page_number}.xlsx"
	unsent_companies_dataframe.to_excel(filename, index = False)

	# OPEN EXCEL WITH COMPANY DETAILS AS LIST OF DICTIONARIES
	unsent_companies_dataframe_read = pd.read_excel(f"clutch_companies_{results_page_number}.xlsx")
	unsent_companies_read = unsent_companies_dataframe_read.to_dict(orient = "records")
	
	# SENDING MESSAGE, UPDATING EXCEL AND TRACKED VARIABLES
	for unsent_company in unsent_companies_read[results_page_company: (len(unsent_companies_read) + 1)]:

		print(f"Working with results_page_company {results_page_company}")
		print(f"Working with total_company_number {total_company_number}")
		send_email_process(unsent_company, driver, email_text)
		total_company_number += 1
		results_page_company +=1
		if results_page_company > 49:
			results_page_company = 0
		update_tracked_variables(total_company_number, results_page_number, results_page_company)
		random_pause(1,4)
	# ADVANCING TO NEXT RESULTS PAGE ONCE ALL COMPANIES IN CURRENT RESULTS PAGE HAVE BEEN SENT AN EMAIL
	results_page_number += 1
	update_tracked_variables(total_company_number, results_page_number, results_page_company)














	
	
	






















