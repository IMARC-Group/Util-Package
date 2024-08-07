import os
import time
import uuid
from util import selenium as sel
import shutil
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains

def test_is_driver_present_firefox():
    print(os.getcwd())
    assert sel._is_driver_present(os.getcwd(), sel.Browser.FIREFOX), "Could not detect geckodriver"


def test_is_driver_present_chrome():
    print(os.getcwd())
    assert sel._is_driver_present(os.getcwd(), sel.Browser.CHROME), "Could not detect chromedriver"


# def init_driver(
#     browser: Browser = Browser.FIREFOX,
#     driver_download_dir: str = None,
#     user_data_dir: str = None,
#     download_dir: str = None,
#     headless: bool = False,
# ) -> webdriver:
def test_init_driver_firefox():
    driver = sel.init_driver(browser=sel.Browser.FIREFOX)
    assert driver is not None, "No webdriver generated."
    driver.quit()


def test_driver_download_dir_firefox():
    download_dir = os.path.join(os.getcwd(), "download_test", str(uuid.uuid4()))
    os.makedirs(download_dir, exist_ok=True)

    driver = sel.init_driver(
        browser=sel.Browser.FIREFOX,
        download_dir=download_dir)

    driver.get("https://library.lol/main/D37C047171EEEF2279D7DC99E9FBF46F")
    driver.find_element(By.XPATH, '//*[@id="download"]/h2[1]/a').click()
    time.sleep(30)
    driver.quit()

    files = os.listdir(download_dir)
    shutil.rmtree(download_dir)

    assert len(files) > 0, "Download directory does not exist."


def test_download_dir_chrome():
    download_dir = os.path.join(os.getcwd(), "download_test", str(uuid.uuid4()))
    os.makedirs(download_dir, exist_ok=True)

    driver = sel.init_driver(
        browser=sel.Browser.CHROME,
        download_dir=download_dir)

    driver.get("https://library.lol/main/D37C047171EEEF2279D7DC99E9FBF46F")
    driver.find_element(By.XPATH, '//*[@id="download"]/h2[1]/a').click()

    time.sleep(30)
    driver.quit()

    files = os.listdir(download_dir)
    shutil.rmtree(download_dir)

    assert len(files) > 0, "Download directory does not exist."


# def test_driver_download_dir_firefox():
    # assert download_dir in driver.capabilities["moz:profile"], "Download directory not set correctly."


def test_user_dir_firefox():
    user_dir = sel.get_profile_path(sel.Browser.FIREFOX)
    driver = sel.init_driver(sel.Browser.FIREFOX, user_data_dir=user_dir)
    print(driver.capabilities)
    assert driver.capabilities["moz:profile"] == user_dir, "profile path not set"


def test_init_driver_chrome():
    driver = sel.init_driver(browser=sel.Browser.CHROME)
    assert driver is not None, "No webdriver generated."
    driver.quit()


def test_init_driver_edge():
    try:
        driver = sel.init_driver(browser=sel.Browser.EDGE)
    except Exception as e:
        assert isinstance(e, NotImplementedError), "No driver generated for Edge browser."


# # TODO: to be completed
# def get_profile_path(browser: Browser) -> str:
#     "returns the path to the profile"
#     profile_path = None
#     return profile_path


def get_profile_firefox():
    path = None
    path = sel.get_profile_path(sel.Browser.FIREFOX)
    assert path is not None, "profile_firefox not found."
    print(path)


def get_profile_chrome():
    path = None
    path = sel.get_profile_path(sel.Browser.CHROME)
    assert path is not None, "profile_firefox not found."
    print(path)
