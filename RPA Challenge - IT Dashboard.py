import re
import time
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.FileSystem import FileSystem
import os
from RPA.PDF import PDF


class Title:

    def __init__(self, base_link):
        self.browser = Selenium()
        self.lib = Files()
        self.pdf = PDF()
        self.url_link = base_link
        self.browser.set_download_directory(os.path.join(os.getcwd(), "output"))
        self.open_browser = self.browser.open_available_browser(base_link)
        self.browser.find_element('//*[@id="node-23"]/div/div/div/div/div/div/div/a').click()
        time.sleep(2)
        self.browser.find_element('//*[@id="agency-tiles-2-widget"]/div/div[9]/div[1]/div/div/div/div[2]/a').click()
        time.sleep(20)
        self.browser.find_element('//*[@id="investments-table-object_length"]/label/select/option[4]').click()
        time.sleep(10)
        self.title = []
        self.budget = []
        self.invest_list = []
        self.links = [{'link': 'https://itdashboard.gov/drupal/summary/422/422-000000004', 'UII': '422-000000004', 'Investment': 'Data Management and Delivery'}, {'link': 'https://itdashboard.gov/drupal/summary/422/422-000001327', 'UII': '422-000001327', 'Investment': 'iTRAK'}, {'link': 'https://itdashboard.gov/drupal/summary/422/422-000001328', 'UII': '422-000001328', 'Investment': 'Mission Support Systems'}]
        self.files = FileSystem()

    def loop_browser(self):

        for ele in self.browser.find_elements('//span[@class="h4 w200"]'):
            self.title.append(ele.text)
            while '' in self.title:
                self.title.remove('')
        for value in self.browser.find_elements('//span[@class=" h1 w900"]'):
            self.budget.append(value.text)
        time.sleep(10)

    def test(self):
        a = self.browser.get_text("id:investments-table-object_info")
        var = a.split(" ")[-2]
        for i in range(1, int(var) + 1):
            UII = self.browser.find_element(f'//*[@id="investments-table-object"]/tbody/tr[{i}]/td[1]').text
            Bureau = self.browser.find_element(f'//*[@id="investments-table-object"]/tbody/tr[{i}]/td[2]').text
            Investment = self.browser.find_element(f'//*[@id="investments-table-object"]/tbody/tr[{i}]/td[3]').text
            Total_Spending = self.browser.find_element(f'//*[@id="investments-table-object"]/tbody/tr[{i}]/td[4]').text
            Type = self.browser.find_element(f'//*[@id="investments-table-object"]/tbody/tr[{i}]/td[5]').text
            CIO_Rating = self.browser.find_element(f'//*[@id="investments-table-object"]/tbody/tr[{i}]/td[6]').text
            Projects = self.browser.find_element(f'//*[@id="investments-table-object"]/tbody/tr[{i}]/td[7]').text
            if UII and Bureau and Investment and Total_Spending and Type and CIO_Rating and Projects:
                self.invest_list.append({"UII": UII, "Bureau": Bureau, "Investment": Investment,
                                         "Total Spending": Total_Spending, "Type": Type, "CIO Rating": CIO_Rating,
                                         "Projects": Projects})
            try:
                link = self.browser.find_element(
                    f'//*[@id="investments-table-object"]/tbody/tr[{i}]/td[1]').find_element_by_tag_name(
                    "a").get_attribute("href")
                self.links.append({"link": link, "UII": UII, "Investment": Investment})
            except:
                pass

    def pdf_download(self, links):
        for link in links:
            self.browser.go_to(link["link"])
            self.browser.wait_until_page_contains_element('//*[@id="business-case-pdf"]/a')
            self.browser.find_element('//*[@id="business-case-pdf"]/a').click()
            while self.files.does_file_not_exist(f"output/{link['UII']}.pdf"):
                pass

    def open_pdf(self):
        for link in self.links:
            try:
                file = f'output/{link["UII"]}.pdf'
                text = self.pdf.get_text_from_pdf(file)
                string = re.split(r"Bureau:|Section B", text[1])[1]
                if link["UII"] in string and link["Investment"] in string:
                    print("The UII and Investment are matched")
                else:
                    print("No Matched")
            except:
                pass

    def workbook(self):
        file = self.lib.create_workbook("output/workbook.xlsx")
        data = {"Title": self.title, "budget": self.budget}
        file.append_worksheet("Sheet", content=data, header=True)
        file.save()
        file.create_worksheet("Individual Investments")
        file.append_worksheet("Individual Investments", content=self.invest_list, header=True)
        file.save()

title_obj = Title("https://itdashboard.gov")
title_obj.loop_browser()
title_obj.test()
title_obj.pdf_download(title_obj.links)
title_obj.open_pdf()
title_obj.workbook()
