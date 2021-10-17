import os
import pandas as pd
from selenium.webdriver.chrome.options import Options
import selenium.webdriver as webdriver
import time
from email.message import EmailMessage
import smtplib
from conf import query, num_page, receiver, login, password

query_link = f"https://www.semanticscholar.org/search?q={query}&sort=relevance&page="
links_list = [query_link + str(i+1) for i in range(num_page)]  # create links to follow



# working paths
working_dir = os.path.dirname(os.path.realpath(__file__))
folder_for_pdf = os.path.join(working_dir, "articles")
webdriver_path = os.path.join(working_dir, "chromedriver")  # proper version https://chromedriver.chromium.org/

# chek if articles directory is exist and create if not
if not os.path.isdir(folder_for_pdf):
    os.mkdir(folder_for_pdf)

# webdriver
chrome_options = Options()
prefs = {"download.default_directory": folder_for_pdf, "download.prompt_for_download": True,
         "download.extensions_to_open": "applications/pdf", "enabled": True, "name": "Chrome PDF Viewer"}
chrome_options.add_experimental_option('prefs', prefs)
os.environ["webdriver.chrome.driver"] = webdriver_path  # 'webdriver' executable needs to be in PATH. Please see https://sites.google.com/a/chromium.org/chromedriver/home

download_dir = "C:\\Users\\omprakashpk\\Documents" # for linux/*nix, download_dir="/usr/Public"
options = webdriver.ChromeOptions()

driver = webdriver.Chrome(executable_path=webdriver_path, options=chrome_options)

final_info = []  # empty dictionary for articles info
for search_link in links_list:
    # get all links to articles from the page
    driver.get(search_link)

    print(search_link)
    time.sleep(5)
    articles = driver.find_elements_by_css_selector(".cl-paper-row.serp-papers__paper-row.paper-row-normal")


    articles_links = []

    for article in articles:
        try:
            link = article.find_element_by_css_selector(
                "a").get_attribute("href")
            articles_links.append(link)
        except:
            print(4)
            pass
    print(9)
    for link in articles_links:
        # get info of each article 
        tmp_info = {}
        driver.get(link)
        text = driver.find_element_by_class_name("flex-item.flex-item--width-66.flex-item__left-column").text
        citations = driver.find_element_by_class_name("paper-detail-page__paper-nav").text
        tmp_info.update({
            'authors': text.split("\n")[2],
            'title': text.split("\n")[1],
            'source': text.split("\n")[0].split("Corpus")[0],
            'description': text.split("\n")[6],
            'number of citations': citations.split("\n")[3].split(" ")[0]
        })
        print(tmp_info)
        print("PDF", text.split("\n")[7])

        # trying to download the article's doc
        if text.split("\n")[7] == "View PDF":
            initial_dir = os.listdir(folder_for_pdf)
            link_to_download = driver.find_element_by_css_selector("a.icon-button.button--full-width.button--primary.flex-paper-actions__button.flex-paper-actions__button--primary").get_attribute("href")
            driver.get(link_to_download)
            print("link_to_download", link_to_download)
            try:

                driver.find_element_by_css_selector("cr-icon-button#download").click()
                time.sleep(5)

                current_dir = os.listdir(folder_for_pdf)
                filename = list(set(current_dir) - set(initial_dir))[0]
                full_path = os.path.join(folder_for_pdf, filename)

            except Exception as e:
                full_path = None
            tmp_info.update({'path_to_file': full_path})
        final_info.append(tmp_info.copy())
        time.sleep(2)
driver.quit()

# write all info to excel
df = pd.DataFrame(final_info)
excel_path = os.path.join(working_dir, "data.xlsx")
df.to_excel(excel_path, index=False)

# create email
mail = EmailMessage()
mail['From'] = login
mail['To'] = receiver
mail['Subject'] = "Topics analysis"
mail.set_content("Hi!\n\nFind attached excel file with articles info.\n\nRegard,")

# add attachment
with open(excel_path, 'rb') as f:
    file_data = f.read()
    file_name = f'articles_info.xlsx'
mail.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

server = smtplib.SMTP('smtp.office365.com', 587)
server.starttls()
server.login(login, password)
server.send_message(mail)
server.quit()


"""



# send email


"""