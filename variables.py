driver_file = 'chromedriver.exe'
excel_file = 'values.xlsx'
url_bnr = "https://www.bnr.ro/Home.aspx"
index = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
line = '1'
xpath_home = "//a[@id='hlHome']"
xpath_dropdown = "// a[contains(text(), 'Politică monetară')]"
xpath_element_dd = "//a[@title='Instrumentele de politică monetară']"
xpath_element_dd_title = "//h1[contains(text(),'Instrumentele de politică monetară')]"
xpath_bnr = "//a[normalize-space()='BNR']"
xpath_element_menu = "//ul[@class='tree']//a[contains(text(),'Comunicate de presă')]"
xpath_calendar = "//a[contains(text(),'Calendarul comunicatelor de presă')]"
xpath_data = "//a[@title='24 mai']"
xpath_data_eveniment = "//span[@id='ctl00_ctl00_CPH1_CPH1_ctl01_lblDataText']"
result_passed = "Passed"
result_failed = "Failed"
