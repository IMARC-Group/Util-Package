from util import selenium as sel

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


