import threading
import time
from datetime import date
from tkinter import scrolledtext, messagebox, Entry, Label, Button, WORD, Tk, END
import pandas as pd
from selenium import webdriver

today = date.today()
d1 = today.strftime("%d_%m_%Y")


def is_not_blank(val):
    if val and val.strip():
        return True
    return False


class PortfolioVisualizerGUI:

    def __init__(self, root):
        root.geometry("500x500")

        root.title('Portfolio Visualizer')

        label_1 = Label(root, text="Timing model:", width=20, font=("bold", 10))
        label_1.place(x=80, y=30)

        self.timing_model = Entry(root)
        self.timing_model.place(x=240, y=30)

        label_3 = Label(root, text="Ticker:", width=20, font=("bold", 10))
        label_3.place(x=68, y=80)

        self.ticker = Entry(root)
        self.ticker.place(x=240, y=80)

        label_4 = Label(root, text="Out of market asset:", width=20, font=("bold", 10))
        label_4.place(x=70, y=130)
        self.asset = Entry(root)
        self.asset.place(x=240, y=130)

        label_5 = Label(root, text="Timing period:", width=20, font=("bold", 10))
        label_5.place(x=70, y=180)
        self.period = Entry(root)
        self.period.place(x=240, y=180)

        label_6 = Label(root, text="Trading frequency:", width=20, font=('bold', 10))
        label_6.place(x=75, y=230)
        self.frequency = Entry(root)
        self.frequency.place(x=240, y=230)

        self.button = Button(root, text='Run', width=20, bg="black", fg='white', command=self.start_on)
        self.button.place(x=180, y=280)

        self.text_area = scrolledtext.ScrolledText(root,
                                                   wrap=WORD,
                                                   width=45,
                                                   height=8,
                                                   font=("Courier", 11))

        self.text_area.place(x=40, y=320)

    def start_on(self):
        t = threading.Thread(target=self.run_on_click)
        t.start()

    def run_on_click(self):
        if self.validate_form():
            self.scrape_data()
        else:
            messagebox.showerror("showerror", "Make sure to fill all the fields")

    def validate_form(self):
        if is_not_blank(self.ticker.get()) and is_not_blank(self.timing_model.get()) and is_not_blank(self.asset.get()) \
                and is_not_blank(self.frequency.get()) and is_not_blank(self.frequency.get()):
            return True
        return False

    def scrape_data(self):
        if self.button["state"] == "normal":
            self.button["state"] = "disabled"
            # object.after(5000, partial(self.scrape_data, self.button))
        else:
            self.button["state"] = "active"
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument("--log-level=3")
        prefs = {"profile.managed_default_content_settings.images": 2}
        chrome_options.add_experimental_option("prefs", prefs)
        driver = webdriver.Chrome(options=chrome_options, executable_path=".\chromedriver.exe")
        driver.implicitly_wait(20)
        all_data = []
        combo = []
        timing_model = self.timing_model.get().split(',')
        out = self.asset.get().split(',')
        timing_period = self.frequency.get().split(',')
        trading_frequency = self.frequency.get().split(',')
        keywords = self.ticker.get().upper().split(',')
        for i in timing_model:
            for j in out:
                for k in timing_period:
                    for l in trading_frequency:
                        data = []
                        data.append(i)
                        data.append(j)
                        data.append(k)
                        data.append(l)
                        combo.append(data)

        c = 0
        k = 0
        old = pd.read_excel('data.xlsx')
        writer = pd.ExcelWriter('data.xlsx', engine='xlsxwriter')
        for key in keywords:
            k += 1
            print('\n\n\nkey :', key)
            self.text_area.insert(END, "\n" + '\nkey :' + key)
            print('{} out of {}'.format(k, len(keywords)))
            self.text_area.insert(END, "\n" + '{} out of {}'.format(k, len(keywords)))
            for com in combo:
                try:
                    present_combo = com
                    c += 1
                    driver.get(
                        'https://www.portfoliovisualizer.com/test-market-timing-model'
                    )
                    # timing model
                    try:
                        time.sleep(1)
                        driver.find_element_by_id('timingModel_chosen').click()
                        time.sleep(1)
                        driver.find_element_by_xpath(
                            '//*[@id="timingModel_chosen"]/div/ul/li[{}]'.format(com[0])).click()
                    except:
                        pass

                    # send tikcer
                    try:
                        driver.find_element_by_id('symbol').send_keys(key)
                        try:
                            driver.find_element_by_xpath('/html/body/div[2]/form/div[15]/div/div/span[2]').click()
                            time.sleep(1)
                            driver.find_element_by_xpath(
                                '/html/body/div[3]/div/div/div[2]/form/div[2]/div/span/input').send_keys(key)
                            driver.find_element_by_xpath(
                                '/html/body/div[3]/div/div/div[2]/form/div[2]/div/span/div/div/div[1]').click()
                            driver.find_element_by_xpath('/html/body/div[3]/div/div/div[3]/button[2]').click()
                        except:
                            continue
                    except:
                        driver.find_element_by_id('symbols').send_keys(key)
                    time.sleep(2)

                    # out of asset model

                    try:
                        driver.find_element_by_id('outOfMarketAssetType_chosen').click()
                        time.sleep(1)
                        driver.find_element_by_xpath('//*[@id="outOfMarketAssetType_chosen"]/div/ul/li[2]').click()
                        driver.find_element_by_id('outOfMarketAsset').send_keys(com[1])
                    except:
                        pass
                    # timing period
                    try:
                        driver.find_element_by_id('windowSize_chosen').click()
                        time.sleep(1)
                        driver.find_element_by_xpath(
                            '//*[@id="windowSize_chosen"]/div/ul/li[{}]'.format(com[2])).click()
                    except:
                        pass

                    # trading frequency
                    try:
                        driver.find_element_by_id('rebalancePeriod_chosen').click()
                        time.sleep(1)
                        driver.find_element_by_xpath(
                            '//*[@id="rebalancePeriod_chosen"]/div/ul/li[{}]'.format(com[3])).click()
                    except:
                        pass

                    time.sleep(2)
                    driver.find_element_by_id('submitButton').click()
                except:
                    continue
                try:
                    tab = driver.find_element_by_xpath('/html/body/div[2]/div[9]/div[1]/table/tbody')

                    rows = tab.find_elements_by_tag_name('tr')
                    rows.pop(1)
                    for row in rows:
                        td = row.find_elements_by_tag_name('td')
                        data = []
                        data.append(present_combo[0])
                        data.append(present_combo[1])
                        data.append(present_combo[2])
                        data.append(present_combo[3])
                        data.append(key)
                        for i in td:
                            data.append(i.text)
                        data.pop(5)
                        data.pop(6)
                        if len(data) == 14:
                            data.append(d1)
                            all_data.append(data)
                            print('Added')
                            self.text_area.insert(END, "\n" + "Added")
                        else:
                            print(len(data))
                except:
                    continue

            df = pd.DataFrame(all_data,
                              columns=['Timing model', 'out of market asset', 'timing period', 'trading frequency',
                                       'Key',
                                       'Final Balance', 'CAGR', 'Stdev',
                                       'Best Year', 'Worst Year', 'Max. Drawdown', 'Sharpe Ratio',
                                       'Sortino Ratio', 'US Mkt Correlation', 'Date'])
            df.to_excel(writer, sheet_name='portfolios', index=False)
            print('Saved..')
            self.text_area.insert(END, "\n" + "Saved..")

        df = pd.DataFrame(all_data,
                          columns=['Timing model', 'out of market asset', 'timing period', 'trading frequency', 'Key',
                                   'Final Balance', 'CAGR', 'Stdev',
                                   'Best Year', 'Worst Year', 'Max. Drawdown', 'Sharpe Ratio',
                                   'Sortino Ratio', 'US Mkt Correlation', 'Date'])
        df = df.append(old)
        df.to_excel(writer, sheet_name='portfolios', index=False)
        highest_sorts_keys = []
        keys = {}
        for data in df.values.tolist():
            key = data[4]
            if key in keys:
                keys[key].append(data)
            else:
                keys[key] = [data]

        for key, value in keys.items():
            value.sort(key=sort_ratio, reverse=True)
            highest_sorts_keys.append(value[0])

        df = pd.DataFrame(highest_sorts_keys,
                          columns=['Timing model', 'out of market asset', 'timing period', 'trading frequency', 'Key',
                                   'Final Balance', 'CAGR', 'Stdev',
                                   'Best Year', 'Worst Year', 'Max. Drawdown', 'Sharpe Ratio',
                                   'Sortino Ratio', 'US Mkt Correlation', 'Date'])
        df.to_excel(writer, sheet_name='Sortino Ratio', index=False)
        driver.close()
        driver.quit()
        del driver
        if self.button["state"] == "disabled":
            self.button["state"] = "normal"
        writer.save()


def sort_ratio(data):
    return float(data[12])


if __name__ == '__main__':
    root = Tk()
    my_gui = PortfolioVisualizerGUI(root)
    root.mainloop()
