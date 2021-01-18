import string, random
import pandas as pd
import numpy as np
import time, os


def listify(l):
    if not type(l)==list:
        return [l]
    else: 
        return l


def add_random_id(message):
    return message + f"  Message Id: {''.join(random.choices(string.ascii_letters + string.digits, k=8))}"


def ClickFunc(xpath, *args, **kwaargs):
    def _inner(browser):
        button = browser.find_element_by_xpath(xpath)
        try: 
            button.click()
        except:            
            js_click(browser, button)
    return _inner


def js_click(browser, element):
    browser.execute_script('arguments[0].click()', element)


def img_loaded(browser, xpath):
    captcha = browser.find_element_by_xpath(xpath)
    loaded = browser.execute_script("return arguments[0].complete && "+
                                "typeof arguments[0].naturalWidth != \"undefined\" && "+
                                "arguments[0].naturalWidth > 0", captcha)
    return loaded


def read_df(file):
    if file.endswith('.csv'):
        try:
            return pd.read_csv(file, index_col=None)
        except pd.errors.EmptyDataError:
            return pd.DataFrame()
    else:
        return pd.read_excel(file, index_col=None)


def save_df(file, df):
    if file.endswith('.csv'):
        return df.to_csv(file, index=False)
    else:
        return df.to_excel(file, index=False)


def get_prop(df, i, prop):
    return (str(df[prop][i])).strip() if prop in df.columns else np.NaN


def combine_dicts(data, d, key_pref=''):
    for k, v in d.items():
        k = k.strip(' :').strip()
        if k == "DOB": k = 'Date of Birth'
        if k == 'Name': k = "Candidate's Name"
        k = key_pref + k
        if k not in data.keys():
            data[k] = v.strip()
    return data


def ldir(path):
    return list(path.iterdir())

def zoom_to(browser, pct=90):
    browser.set_window_size(1920, 1280)
    browser.execute_script(f"document.body.style.zoom='{pct}%'")

from PIL import Image

def fullpage_screenshot(driver, file):

        # print("Starting chrome full page screenshot workaround ...")

        total_width = driver.execute_script("return document.body.offsetWidth")

        total_height = driver.execute_script("return document.body.parentNode.scrollHeight")

        viewport_width = driver.execute_script("return document.body.clientWidth")

        viewport_height = driver.execute_script("return window.innerHeight")

        # print("Total: ({0}, {1}), Viewport: ({2},{3})".format(total_width, total_height,viewport_width,viewport_height))

        rectangles = []

        i = 0

        while i < total_height:

            ii = 0

            top_height = i + viewport_height

            if top_height > total_height:

                top_height = total_height

            while ii < total_width:

                top_width = ii + viewport_width

                if top_width > total_width:

                    top_width = total_width

                # print("Appending rectangle ({0},{1},{2},{3})".format(ii, i, top_width, top_height))

                rectangles.append((ii, i, top_width,top_height))

                ii = ii + viewport_width

            i = i + viewport_height

        stitched_image = Image.new('RGB', (total_width, total_height))

        previous = None

        part = 0

        for rectangle in rectangles:

            if not previous is None:

                driver.execute_script("window.scrollTo({0}, {1})".format(rectangle[0], rectangle[1]))

                # print("Scrolled To ({0},{1})".format(rectangle[0], rectangle[1]))

                time.sleep(0.2)

            file_name = "part_{0}.png".format(part)

            # print("Capturing {0} ...".format(file_name))

            driver.get_screenshot_as_file(file_name)

            screenshot = Image.open(file_name)

            if rectangle[1] + viewport_height > total_height:

                offset = (rectangle[0], total_height - viewport_height)

            else:

                offset = (rectangle[0], rectangle[1])

            # print("Adding to stitched image with offset ({0}, {1})".format(offset[0],offset[1]))

            stitched_image.paste(screenshot, offset)

            del screenshot

            os.remove(file_name)

            part = part + 1

            previous = rectangle

        stitched_image.save(file)

        # print("Finishing chrome full page screenshot workaround...")