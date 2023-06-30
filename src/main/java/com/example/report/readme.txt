## 简单域处理
处理 w:fldSimple 元素
<w:fldSimple>需要保留的内容</w:fldSimple>

## 合并域数据
无论是手动将docx另存为xml或者使用spiredoc的命令将docx另存为xml，都会出现域内容被分段的情况
这里需要将一个域中完整的数据拼接到一起
处理 w:fldChar 元素
<w:fldChar w:fldCharType="begin" />
<w:instrText xml:space="preserve" />
<w:fldChar w:fldCharType="separate" />
<w:t> 需要保留的内容 </w:t>
<w:fldChar w:fldCharType="end" />

## velocity一般语法
处理在独立段落中，使用velocity的set、if、foreach、macro等语法
前提是都是独立的段落中，即语法所在的父元素wp是后面要被移除的

## 表格行合并
#tbl_vmerge(val)
参数为是否开始行合并，
true: 则在w:tc的w:tcPr中加入<w:vMerge w:val="restart"/>
false: 在w:tc的w:tcPr中加入<w:vMerge/>

## 表格合并列
#tbl_grid_span(val)
参数为跨越的列数，注意：参数必须是字符串类型
如果列数大于1，则添加w:gridSpan元素，并配置w:val的值
如果列数为1，则不添加w:gridSpan
如果列数为0，则删除当前的tc

## 占位图片
处理占位图片元素，
0，在word中用”替换文字“用域来标记图片，
1，在xml中通过 w:drawing 标签下 wp:docPr 的 descr 属性的值来找到对应的域
2，找到对应的a:blip 标签的 r:embed 属性值，就是word给图片分配的id
3，通过id属性找到对应的Relationship，获取Target属性的值，就是对应图片在word中的实际名称
4，通过名称（前面要拼上/word/）找到对应图片的base64编码，将编码内容替换为域名
5，将前面 wp:docPr 的descr属性删除，将pic:cNvPr 的descr属性删除

## foreach图片
#pic_foreach, #pic_end 处理图片foreach循环
思路：找到#pic_foreach, #pic_end标签，紧跟在#pic_foreach所在的w:p之后的w:p兄弟元素就是图片元素
替换图片元素中的rId内容，引入$!{foreach.index}实现图片数量的动态变化
新增Relationship及pkg:part，同样引入$!{foreach.index}实现图片数量的动态变化