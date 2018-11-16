# exceltoword
java读取excel内容转换为固定样式的word

读取excel数据：
1、引入poi相关jar
2、poi读取excel内容循环遍历封装数据

填充数据到word：
1、修改word中需要用数据进行填充的内容为固定的key。如：test1/test2...
2、把word另存为xml
3、修改xml文件后缀.ftl
4、java封装填充数据如：test1/test2...
5、读取模板.ftl文件并写入数据

注：如果需要word中部分内容多次循环生成，需要将此部分的.ftl代码用下面循环体处理
<#list list as test>需要循环的ftl部分代码（内部填充字段${test.xxx}）</#list>

参考文档：
https://www.cnblogs.com/lcngu/p/5247179.html
