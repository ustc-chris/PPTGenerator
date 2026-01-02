# PPTGenerator


# 建议：
（USTChorus模版）
先去USTChorus目录下，查看All.txt文件
然后运行1.1.SidePPT.py，查看结果
然后运行1.2.SrcPPT.py，查看结果
（进阶使用）
先去example目录下，运行1.process_to_json.py，然后在example/output目录下，查看输出的All.json，了解json格式
然后查看2.json_to_ppt.py，了解如何从json格式文件创建PPT，并且运行，查看结果
然后查看3.ppt_to_json.py，了解如何从PPT文件提取为json文件，并且运行，查看结果



## 用法：
在`share/Template.py`文件中，定义了一个Manager类，用于管理模板。你可以从现有模板中选取，也可以自己定义新的模板。

当你定义新的模板的时候，需要
0. 在`./share/template.pptx`文件中，打开母版模式，新建一个板式，并且（不要用WPS，用Microsoft Powerpoint）增加你需要的占位符，调整大小格式等。
0. 运行一次`script/Check_template.py`，记录你的新的板式中的占位符对应的id
1. 在TemplateType中增加这个一个枚举，值为这个id
2. 在下面增加你要的这个类，继承自Template类，实现各个方法
    - 其中parse方法是用来解读json格式文件的（详见`Manager.parse()`方法）；
    - print方法是用来输入到最终的PPT里的（详见`Manager.print()`方法）；
    - convert_back方法是用来从PPT里提取数据的（详见`Manager.convert_back()`方法）。

你可以用自己的任意方式来解析文本文件，然后创建json格式的文件，用于输入到PPTGenerator中。详见示例`example/*.py`