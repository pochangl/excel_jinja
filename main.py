from jinja2 import Environment
from tempfile import TemporaryDirectory
import zipfile
import os
import shutil

xlsx_file = 'example.xlsx'
output1 = 'example1.xlsx'
data1 = {'a': '變數 1', 'b': '變數 2'}

output2 = 'example2.xlsx'
data2 = {'a': 'var 1', 'b': 'var 2'}


env = Environment()

with TemporaryDirectory() as dir:

    with zipfile.ZipFile(xlsx_file) as zip:
        zip.extractall(dir)

    # 讀取 template
    shared_strings_file = os.path.join(dir, 'xl/sharedStrings.xml')
    with open(shared_strings_file) as ssf:
        template = env.from_string(ssf.read())

    # 替換 sharedStrings.xml
    with open(shared_strings_file, 'w') as ssf:
        ssf.write(template.render(data1))

    # 生成檔案 output1.xlsx
    shutil.make_archive(output1, 'zip', dir)
    os.rename(f'{output1}.zip', output1)


    # 替換 sharedStrings.xml
    with open(shared_strings_file, 'w') as ssf:
        ssf.write(template.render(data2))

    # 生成檔案 output2.xlsx
    shutil.make_archive(output2, 'zip', dir)
    os.rename(f'{output2}.zip', output2)