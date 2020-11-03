# git

## 1.git 的远程操作（以github为例）

### 1.1 下载远程仓库的代码

首先需要在本地初始化一个 git 仓库

使用 git init 命令

初始化后，可以在本地对应文件夹中看到后缀为 .git 的隐藏文件夹

其次使用 git clone XXX 命令拉取代码，此处的 XXX 应该为对应远程库的链接

例如 git clone https://github.com/cyk2018/word.git

### 1.2 让本地库链接远程库

本地库链接远程库之后，才能使用 pull push 等命令

连接命令 git remote add XXX YYY, 其中 XXX 为链接名, YYY 为链接

例如 git remote add origin https://github.com/cyk2018/word.git

## 2. git 的本地操作

git 本地的三块区域（工作区、暂存区、提交区）

因此写完代码之后，需要从工作区到暂存区再到提交区。

git add * 表示将当前文件夹中一切内容暂存

git commit -m "阿巴阿巴" 表示将暂存区内容提交到提交区，阿巴阿巴为提交信息。

git pull origin master:master

git push origin master:master

这两个命令，先从远程拉取最新代码并合并，然后推送到远程仓库

其中的 origin 可以替换为各自的远程连接名。



## 3、也可以使用图形化界面工具

自己搜索吧



