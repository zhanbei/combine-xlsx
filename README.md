# 将多个 XLSX 合并并一个 XLSX

<!-- > 2019-04-19T11:31:38+0800 -->

此应用程序功能是把读取文件夹内的所有 XLSX 文件并 **按照指定的特定列** 合并为一个 CSV 文件，然后根据需要可由外部工具将 CSV 文件转化为需要的 XLSX 文件。

此应用程序使用 NodeJs 实现，所以运行脚本需要 NodeJs 语言环境支持。

暂无排序功能。

另一种做法是读取所有 XLSX 文件，去除第一行，然后拼接成一个完整的 CSV/XLSX 文件，此程序目前不是这种做法。

## 使用环境准备

- [x] 安装 [NodeJs](https://nodejs.org/) 软件

## 程序准备

根据说明与实际情况修改 `index.js` 文件初始的 `#mTitleKeys` 和 `#mTargetCsvTitles` 定义并保存脚本。

## 命令行中使用

通过 CMD 进入命令行，或者安装 [Git Bash](https://git-scm.com/) 并进入命令行，或者进入 Linux 下的 Terminal/终端。

运行命令参考以下：

```bash
# 合并当前文件夹内的 XLSX 文件，并输出结果到当前文件夹。
node combine-xlsx/index.js 目标文件夹名字

# 合并目标文件夹内的 XLSX 文件，并输出结果到当前文件夹。
node combine-xlsx/index.js 目标文件夹名字
```

## 后续操作步骤

打开生成的 CSV 文件，打开验证数据是否正确，根据需要可另存为 XLSX 文件。

可以尝试使用 Excel 中的排序功能对数据排序。

删除生成的(用于调试的)文件 `output.*.json` (如 `output.raw.json`, `output.parsed.json`, `output.mixed.json`)。
