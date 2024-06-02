# ExcelSpotLight
将选中的单元格所在的行和列涂成淡黄色

## 配置方法
### 创建PERSONAL.XLSB
#### 如果在VBE的项目中没有发现PERSONAL.XLSB文件,可以按以下方法设置:
1. 在开发工具选项卡中选择录制宏,将保存在选项设置为:个人用工作簿
2. 随便录制一些宏
3. 打开VBE就可以看见PERSONAL.XLSB文件了
这个文件是每次打开Excel就会自动加载,所以设置完成后就可以在任意一个Excel中使用了~

### 导入文件
- 将ThisWorkbook.cls 导入到PERSONAL.XLSB中.或者将该文件中第10行及之后的代码手动粘贴至PERSONAL.XLSB的ThisWorkbook中.
- 将ControlSpotLight.bas 导入到PERSONAL.XLSB中.
- 将AppEvents.cls 导入到PERSONAL.XLSB中.

### 便捷性设置
- 导入完成后就可以开始使用了,在使用之前需要将Excel文件全部关闭,重新打开就可以正常加载模块了.
- Excel最上方的栏是快速访问工具栏.可以进行自定义设置.将聚光灯的开关设置在快速访问工具栏内,就可以随时切换聚光灯的开关了~
  - 右键快速访问工具栏,点击自定义快速访问工具栏.
  - 在`从下列位置选择命令`中选择:宏
  - 点击PERSONAL.XLSB!ToggleSpotlight
  - 选择添加
  - 在右边点击该宏,还可以自定义名称和图标.

## 使用方法
默认打开Excel时关闭聚光灯,点击快速访问工具栏内的开关将聚光灯打开.
此时点击任意单元格,就可以看见聚光灯效果.
再次点击快速访问工具栏内的聚光灯开关即可关闭聚光灯.
