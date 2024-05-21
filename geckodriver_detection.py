import os
import platform
from zipfile import ZipFile
import requests


def is_gecko_driver_present(gecko_url):

    current_dir = os.getcwd()
    gecko_path = os.path.join(current_dir, "geckodriver.exe")

    if os.path.isfile(gecko_path):
        return True
    
    return False


def download_gecko_driver(gecko_urls, driver_dir):
    platform_info = platform.architecture()

    gecko_url = None

    match platform_info:
        case ('64bit', 'WindowsPE'):
            gecko_url = gecko_urls[platform_info[0]]
        
        case ('32bit', 'WindowsPE'):
            gecko_url = gecko_urls[platform_info[0]]

        case _:
            raise Exception("Unknown Platform Found.")

    filename = os.path.basename(gecko_url)

    response = requests.get(gecko_url)

    with open(filename, 'wb') as file:
        file.write(response.content)

    with ZipFile(filename, 'r') as zip_ref:
        zip_ref.extractall(driver_dir)

    try:
        os.remove(filename)
    except FileNotFoundError:
        print("Warning: Something might be wrong with availability of the driver. Please check once.")
        pass


def verify_driver(driver_dir: str | os.PathLike):

    gecko_urls = {
        "32bit": "https://github.com/mozilla/geckodriver/releases/download/v0.34.0/geckodriver-v0.34.0-win32.zip",
        "64bit": "https://github.com/mozilla/geckodriver/releases/download/v0.34.0/geckodriver-v0.34.0-win64.zip",
    }

    if not is_gecko_driver_present(driver_dir):
        print('Geckodriver not present. Downloading...')
        download_gecko_driver(gecko_urls, driver_dir)
    else:
        print('Geckodriver present.')


# print(gecko_path)
# gecko_path = get_gecko_driver_path(gecko_url)

if __name__ == '__main__':
    driver_dir = os.getcwd()
    verify_driver(driver_dir)
