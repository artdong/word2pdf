# word2pdf

一个基于python实现的word转pdf文档的工具

## 使用方法

```
git clone https://github.com/artdong/word2pdf.git

cd word2pdf

python3 -m venv venv

source venv/bin/activate

pip3 install -r requirements.txt

python main.py
```

## 配置修改

配置文件config.cfg

doc_folder：指定存放word文件的文件夹，默认doc

pdf_folder：指定存放pdf文件的文件夹，默认pdf

max_worker：同时工作的进程数，默认10
