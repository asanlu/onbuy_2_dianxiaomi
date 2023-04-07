import logging

# 配置logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(filename)s:[line:%(lineno)d] - %(levelname)s: %(message)s',
    # datefmt='%Y-%m-%d %H:%M:%S %p',
    handlers=[
        # for logs write in file（mode：a为追加log，设置为w则表示每次清空，重新记录log）
        logging.FileHandler(f"./onbuy.log", mode="a"),
        logging.StreamHandler()  # for print at console
    ]
)
