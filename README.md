# 音击B45生成器

公开仓库：

## 运行环境

确保计算机上安装了python，并安装以下依赖：

```
pip install selenium urllib sqlite3 pyyaml openpyxl pillow
```

## 使用方式

### 快速使用

随release包附带发行时的全曲目数据，在score.xlsx中填写成绩后，执行：
```
python -u gen_b45.py
```
即可获得名称为'b45.png'的成绩图

### 详细使用说明

0. 配置congfig.yaml配置文件，各参数作用见注释
1. 执行以下指令抓取网站数据（全量模式下较慢）
   ```
   python -u get_all_songs.py
   ```
2. 执行以下指令生成成绩表格，该过程会从已有成绩表中继承数据
   ```
   python -u gen_score_table.py
   ```
3. 填写成绩表
4. 执行以下指令生成B45成绩图
   ```
   python -u gen_b45.py
   ```