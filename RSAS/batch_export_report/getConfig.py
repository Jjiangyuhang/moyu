import configparser, os

# 读取配置文件
def getConfig(filename, section, option):
    """
    :param filename 文件名称
    :param section: 服务
    :param option: 配置参数
    :return:返回配置信息
    """
    # 获取当前目录路径
    proDir = os.path.split(os.path.realpath(__file__))[0]
    # print(proDir)

    # 拼接路径获取完整路径
    configPath = os.path.join(proDir, filename)
    # print(configPath)

    # 创建ConfigParser对象
    conf = configparser.ConfigParser()

    # 读取文件内容
    conf.read(configPath)
    config = conf.get(section, option)
    return config

aesKey = getConfig('config.ini', 'config', 'aeskey').encode()
username = getConfig('config.ini', 'config', 'username').encode()
password = getConfig('config.ini', 'config', 'password').encode()
baseURL = getConfig('config.ini', 'config', 'baseURL').encode()
baseIP = getConfig('config.ini', 'config', 'baseIP').encode()
