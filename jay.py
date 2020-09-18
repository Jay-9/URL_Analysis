import os
import re


# # os.system('chcp 65001')
# os.chdir('C:/d')
# with os.popen(r'dig @211.141.0.99 %s' % 'www.baidu.com', 'r') as f:
#     text = f.read()
#     print(text)
#     match_aliases = re.findall(r'ANSWER SECTION:[\s\S]*', text)
#     print('*'*11)
#     print(match_aliases)


# x = re.findall(r'^ll$ ', 'hello world')
# print(x)


x = re.findall("^hello$", "hello world")
print(x)
