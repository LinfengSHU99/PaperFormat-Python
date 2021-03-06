## 需要的功能：
&emsp;&emsp;图和表的标题序号是否与章节序号一致。图和表是否在正文中有相应的引用。
## 2021.11.14
### 更新的功能：
&emsp;&emsp;完成对目录的字符串列表的提取，可以用来辅助判断一级标题。  
&emsp;&emsp;对于isContent函数增加了一些判断条件，使得判断一个段落是否属于目录更加准确。  
### 待解决的问题：
&emsp;&emsp;关于tab与段落缩进之间的关系。  
&emsp;&emsp;对于某些具有数值属性的格式，虽然其数值与要求的数值不同，但在视觉效果上与要求几乎没有区别。对于此类问题可以考虑设置一个范围，只要格式的属性数值在范围之内的都算正确。
## 2021.11.15
### 更新的功能：
&emsp;&emsp;修复一个bug，该bug曾导致某些有超链接的标题不能准确的定位其父\<w:p>节点，进而不能准确识别其字格式。  
&emsp;&emsp;初步增加了判断一个段落是否为表格标题或者图标题的功能
&emsp;&emsp;将致谢，参考文献，附录加入到标题判断中。
## 2021.11.19
### 更新的功能：
&emsp;&emsp;针对使用word自动编号的标题，仿照大连理工的程序，增加了对该种标题的判断功能（使用与目录各个段落的最小编辑距离辅助判断）
### 待解决的问题：
~~&emsp;&emsp;需要增加对使用word自动编号的标题的标题类别判断功能~~
~~在王凯毕业论文中有一些混合使用手动编号和自动编号的四级标题，需要更完善的考虑。~~

## 2021.11.21
### Update:
&emsp;&emsp;Add some conditional sentence to decide which type the title with auto numbering
belongs to. Fix a bug that leads to the failure of recognizing the content part.

## 2021.11.23
### 更新的功能：
&emsp;&emsp;提取出正文里有边框的表格（避免检测一些作者本意不是表格的表格）。对这些表格完成了是否有标题，标题位置，标题格式的匹配并输出错误信息。完成对表格是否为三线表的匹配，输出错误信息。
### 待解决的问题：
&emsp;&emsp;对表格标题的编号，~~是否有标点进行检测。~~

## 2021.11.26
### 更新的功能：
&emsp;&emsp;完成对表格标题是否有标点，表格正文内容的字体字号检测。

## 2021.12.7
### 更新的功能：
&emsp;&emsp;完成对图是否有标题，是否单独成段，标题位置，标题字体字号，是否有标点，图宽度是否超过页边距的检测，并输出错误信息。  
### 待解决的问题
&emsp;&emsp;若图连续出现，则在检测标题时可能不准确（与表格的问题相同）。
&emsp;&emsp;若图连续出现，且在一个段落中，则匹配可能不准。

## 2021.12.27
### 更新的功能：
&emsp;&emsp;完成对run格式具体错误信息的输出，并将底色设置为黄色。对错误信息增加【】。
### 待解决的问题
&emsp;&emsp;图和表的标题序号是否和章节对应，以及正文中是否出现了相应的图和表的编号，而不是“见下表”，“见下图”等等。