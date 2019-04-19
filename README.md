# 将多个 XLSX 合并并一个 XLSX

<!-- > 2019-04-19T11:31:38+0800 -->

## 使用环境准备

- [x] 安装 [NodeJs](https://nodejs.org/) 软件

## 命令行中使用

通过 CMD 进入命令行，或者安装 [Git Bash](https://git-scm.com/) 并进入命令行，或者进入 Linux 下的 Terminal/终端。

运行命令参考以下：

```bash
# 合并当前文件夹内的 XLSX 文件，并输出结果到当前文件夹。
node combine-xlsx/index.js 目标文件夹名字

# 合并目标文件夹内的 XLSX 文件，并输出结果到当前文件夹。
node combine-xlsx/index.js 目标文件夹名字
```

打开生成的 CSV 文件，打开验证数据是否正确，根据需要可另存为 XLSX 文件。

删除生成的(用于调试的)文件 `output.*.json` (如 `output.raw.json`, `output.parsed.json`, `output.mixed.json`)。
