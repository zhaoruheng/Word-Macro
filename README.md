## 一个可以实现仿真手写的宏，支持随机修改字号，行间距，字体，字间距
###### 2022.12.24 updata:优化注释，增加了一些实践经验。感谢GitHub@[jackatlascn](https://github.com/jackatlascn)贡献
###### 2020.11.1  updata:修复部分bug
###### 2020.12.20 updata:优化代码，处理字间距、行间距、字偏移等细节，增加首行缩进。感谢Github@[喃南下](https://github.com/Airomeo)贡献
###### 2020.9.12  updata：增加随机字间距功能，加了一种推荐字体
![Word-Macro](https://socialify.git.ci/zhaoruheng/Word-Macro/image?font=Inter&forks=1&issues=1&owner=1&pattern=Signal&pulls=1&stargazers=1&theme=Light)
---
#### 以下教程感谢知乎@[Canscxs](https://www.zhihu.com/people/cans-18-32)贡献，如果图片加载有问题请移步[知乎](https://zhuanlan.zhihu.com/p/338196683)

首先单击**视图**>**宏**>**查看宏**(Mac用户为*工具*>*宏*)

![1](https://github.com/zhaoruheng/Word-Macro/blob/master/image/1.jpg?raw=true)

给宏命名后，点击**创建**

![2](https://github.com/zhaoruheng/Word-Macro/blob/master/image/2.jpg?raw=true)

然后复制*代码1.1*文件中代码，粘贴在Sub和End Sub之间，然后按下ctrl+s，关闭界面即可。

ps:默认参数可能并不适合你的需求，代码中提供了各种参数的修改指引，你可以根据指引调出需要的效果

![3](https://github.com/zhaoruheng/Word-Macro/blob/master/image/3.jpg?raw=true)

可以根据自己的需求改字体，前提是**你的电脑上有相同名称的字体**，下面会推荐几个。

![4](https://github.com/zhaoruheng/Word-Macro/blob/master/image/4.jpg?raw=true)

好了之后再点运行

![5](https://github.com/zhaoruheng/Word-Macro/blob/master/image/5.jpg?raw=true)

下面是运行的效果，自己调一下参数会更好
![6](https://github.com/zhaoruheng/Word-Macro/blob/master/image/6.png?raw=true)
![7](https://github.com/zhaoruheng/Word-Macro/blob/master/image/7.png?raw=true)

下面是旧版代码运行效果，如有需要可以从*代码1.0*中复制

![](https://s1.ax1x.com/2020/08/08/a5NtP0.jpg)

---
#### 一些实践经验：

*   GitHub@[jackatlascn](https://github.com/jackatlascn)：
      
      如果仅仅是为了应付简单的检查，随便用一款手写字体即可，如果需要模仿自己的字迹，我的经验可能有点用处。

      1.寻找大量手写字体，我找了100多款。全部安装

      2.下载安装【字由】这个软件，这个软件可以输入一排字，比较不同字体呈现的效果。当然这个软件的付费功能也可以帮你寻找合适的字体。我比较习惯行楷。正楷和行草属实没法用。

      3.人工找合适自己的字体，记下名字。你以为这就结束了吗？这远没有结束。手写字体五花八门，有的是钢笔，有的是粗水笔，还要进一步挑选。

      4.在Excel里输入一排示范字体，最好是一句完整的话，然后每一行用不同字体，字号18-20号，确定一下哪个字号适合自己。

      5.字号确定后，再试一遍打印效果，确定一下哪个字体的打印效果跟你手写的粗细差不多。

      6.大概率，你会发现全都不合适。这里有我的血泪经验分享一下，找一下打印机设置，有没有自动帮你加黑加粗的选项。我取消之后，便找到了3款适合自己粗细、且风格和我相似的字体。到这，基本上大功告成了。接下来就是找一支特别黑的笔，因为激光打印机效果特别黑，正常五毛一支的晨光不太行。但听说高精度的喷墨打印机效果刚刚好，我试了低精度喷墨打印机，字体细节不太行。但身边没有高精度的做实验。

      再点评一下经历，最难的是3的这一步，估计可以找到20多种貌似合适的字体，需要花费1个小时慢慢看。

---

下面推荐的字体有

>汉仪晨妹子W

>汉仪平安行粗简(不太推荐)

>杨任东竹石体-Extralight

>凌慧体-简（macOS自带）

>萌妹子体

>张维镜手写楷书

>手写大象体

>陈静的字完整版

可以在以下链接中找到：

[字客网](https://www.fontke.com/)
[猫啃网](https://www.maoken.com/)
[字体下载](https://www.qiuziti.com/)
[字由](http://www.hellofont.cn)
---
#### 如果你想使用自己的字体，[这个网站](http://59.108.48.27/flexifont-chn/login/)可以帮你制作你自己的手写体，非常简单，直接按网站提示来就行
---
#### 如果这个项目对你有帮助，请点一下右上方star（收藏）
