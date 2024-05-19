import webbrowser
import sys
def open_urls(file_path):
    with open(file_path, 'r') as file:
        urls = file.readlines()

    for url in urls:
        url = url.strip()
        webbrowser.open(url)

if __name__ == '__main__':
    path = sys.argv[1]
    open_urls(path)