# Image 指令说明

**Image** 指令用于在 Excel 报表中输出图像。

### 指令的使用

在 [jxls-demo](https://bitbucket.org/leonate/jxls-demo) 示例项目查看 **org.jxls.demo.ImageDemo** 中 **Image** 指令的操作，该示例演示了如何使用 Java API 定义 **Image** 指令。

如果您希望使用 Excel 标记来定义命令，如下：

```text
jx:image(lastCell="D10" src="image" imageType="PNG")
```

以上的标记中，**lastCell** 定义了包含区域的图像的右下角单元格。如果注释被放置在单元格 **A1** 中，那么图像将被放置在输出 Excel 中的 **A1:D10** 区域中。
**src** 指包含图像字节的 jxls 上下文中的变量名称。
**imageType** 定义了图像类型，可以是以下类型之一: **PNG**、**JPEG**、**EMF**、**WMF**、**PICT**、**DIB** 。**imageType** 属性的默认值是 **PNG** 。在上面的例子中，我们可以跳过它

Java 代码应该包含类似这样的内容，以便将所需的图像放到 image 属性下的上下文中

```java
InputStream imageInputStream = ImageDemo.class.getResourceAsStream("business.png");
byte[] imageBytes = Util.toByteArray(imageInputStream);
context.putVar("image", imageBytes);
```

> [原文地址](http://jxls.sourceforge.net/reference/image_command.html)
