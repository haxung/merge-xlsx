# merge-xlsx

定制化工具，用于合并和格式化 `excel` 文件。

## 参数说明

|      参数      |           说明           |           默认值           |  是否必传  |
|:------------:|:----------------------:|:-----------------------:|:------:|
|  -h, --help  |        显示工具参数信息        |         ------          | ------ |
| -f, --folder |   需要合并的 excel 文件夹路径    |   D:/work/week-report   |   否    |
| -o, --order  |    合并 excel 时使用的顺序     |        林鲁单冀坤朱茗马涂        |   否    |
|  -x, --xlsx  |   合并后生成的 excel 文件路径    | D:/work/week-report/merge.xlsx |   否    |
|  -p, --pdf   | excel 转 pdf 时 pdf 文件路径 | D:/work/week-report/merge.pdf  |   否    |
|  -v, --vba   |      生成的 vba 文件路径      | D:/work/week-report/merge.vba  |   否    |
|  -w, --week  |      表头使用的时间日期范围       |    北京时间现在所在星期的周一到周日     |   否    |
|  -d, --day   |      个人表单使用的时间日期       |      北京时间现在所在星期的周五      |   否    |

## 使用说明

### 环境准备

* 安装 `python 3.9` 以上版本
* 打开 `merge-xlsx` 项目，运行以下命令：

  ```shell
  pip install -r requirements.txt
  ```

### 合并表格

```shell
python xlsx2pdf.py -f D:/work/week-report
```

运行上述命令后会生成 `merge.xlsx, merge.xlsm, merge.vba`，因 `merge.xlsx` 无法直接添加宏，所以需要 `merge.xlsm`。

### 写入宏

1. 打开 `merge.vba`，全选复制。
2. 打开 `merge.xlsm`，点击 `视图->宏`，输入任意宏名，点击右侧的 `创建` 按钮，全选粘贴。
3. 点击 `视图->宏`，此时只有一个宏，宏名为 `打印`，点击右侧的 `运行` 按钮，等待宏命令执行完毕，自动在浏览器打开 `PDF` 文件。