> 提取扫描资产中，有`web`服务的资产，以`http/https://ip:port`形式保存到压缩包所在目录下的`url.txt`中

运行前需要安装`python`库

- bs4

**useage**

```bash
python web_from_RSAS.py 导出的扫描结果html的压缩包绝对路径

eg.
python web_from_RSAS.py D:\work\test\xxxx_2024_02_23_html.zip
```

在浏览器中批量打开`url.txt`文件中的地址

```python
import webbrowser

def open_urls(file_path):
    with open(file_path, 'r') as file:
        urls = file.readlines()

    for url in urls:
        url = url.strip()
        webbrowser.open(url)

if __name__ == '__main__':
    open_urls(r'D:\work\test\urls.txt')
```

