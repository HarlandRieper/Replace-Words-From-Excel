# ReplaceWordsFromExcel word文档的批量“查找和替换”
在word文档中，需要在不改变原格式的前提下，将“苹果”替换为“菠萝”、“香蕉”替换为“榴莲”、“葡萄”替换为“橙子”......手动输入也太累了，试试宏呢？

# 如何导入宏？

- Alt + F11 进入编辑界面
- 在 /Normal/Modules 中找到 NewMarcos 双击打开


![Image](https://user-images.githubusercontent.com/59085287/236612012-484f4fef-842e-47a4-bed7-1526057a8373.png)


- 将代码粘贴在文件末尾
- 保存、关闭

# 如何使用宏？

## 新建 excel 文件
- 在sheet1中第一列写出需要被替换的词，第二列中列出替换后的词。下图为一个示例。

![Image](https://user-images.githubusercontent.com/59085287/236611936-3b91f672-8bf9-4ac4-99d5-a14ab5cffbf6.png)

- 获取该文件地址。win11中可右键该 excel 文件，选择“复制文件地址”





## 开始替换 word 文档中的内容
- word文件中选择“视图”（view）选项卡，点击“宏”（Marcos）。或在搜索栏中搜索。




- 选择“ReplaceWordsFromExcel”，点击“运行”（run）



![Image](https://user-images.githubusercontent.com/59085287/236612056-3098ab4e-99af-4ca5-bc99-711709c025c4.png)


- 弹出输入框，将 excel 文件路径输入进去。不需要考虑斜杠和双引号的问题。
- 弹出选择框，如果需从excel第一列替换为第二列，则选择“是”；若需要反向替换，选择“否”

![image](https://user-images.githubusercontent.com/59085287/236615170-fcf285e0-6445-440f-b582-d218bb11d7b5.png)

- 点击“ok”，替换完毕。

# 便捷使用

可以把常用的宏加入在选项卡中。
