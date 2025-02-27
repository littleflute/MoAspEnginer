﻿var text_input = "文字";
var help_mode = "UBB 代码 - 帮助信息\n\n点击相应的代码按钮即可获得相应的说明和提示";
var adv_mode = "UBB 代码 - 直接插入\n\n点击代码按钮后不出现提示即直接插入相应代码";
var normal_mode = "UBB 代码 - 提示插入\n\n点击代码按钮后出现向导窗口帮助您完成代码插入";
var email_help = "插入邮件地址\n\n插入邮件地址连接。\n例如：\n[email]webmaster@loveyuki.com[/email]\n[email=webmaster@loveyuki.com]Loveyuki[/email]";
var email_normal = "请输入链接显示的文字，如果留空则直接显示邮件地址。";
var email_normal_input = "请输入邮件地址。";
var fontsize_help = "设置字号\n\n将标签所包围的文字设置成指定字号。\n例如：[size=3]文字大小为 3[/size]";
var fontsize_normal = "请输入要设置为指定字号的文字。";
var font_help = "设定字体\n\n将标签所包围的文字设置成指定字体。\n例如：[font=仿宋]字体为仿宋[/font]";
var font_normal = "请输入要设置成指定字体的文字。";
var bold_help = "插入粗体文本\n\n将标签所包围的文本变成粗体。\n例如：[b]Loveyuki's BLOG[/b]";
var bold_normal = "请输入要设置成粗体的文字。";
var page_help = "插入分页\n\n在标签所在位置插入分页标识符。\n例如：[page][/page]";
var page_normal = "插入分页标识符。";
var italicize_help = "插入斜体文本\n\n将标签所包围的文本变成斜体。\n例如：[i]Loveyuki's BLOG[/i]";
var italicize_normal = "请输入要设置成斜体的文字。";
var quote_help = "插入引用\n\n将标签所包围的文本作为引用特殊显示。\n例如：[quote]Loveyuki 版权所有 - Loveyuki's BLOG[/quote]";
var quote_normal = "请输入要作为引用显示的文字。";
var color_help = "定义文本颜色\n\n将标签所包围的文本变为制定颜色。\n例如：[color=red]红颜色[/color]";
var color_normal = "请输入要设置成指定颜色的文字。";
var center_help = "居中对齐\n\n将标签所包围的文本居中对齐显示。\n例如：[align=center]内容居中[/align]";
var center_normal = "请输入要居中对齐的文字。";
var link_help = "插入超级链接\n\n插入一个超级连接。\n例如：\n[url]http://www.loveyuki.com[/url]\n[url=http://www.loveyuki.com]Loveyuki's BLOG[/url]";
var link_normal = "请输入链接显示的文字，如果留空则直接显示链接。";
var link_normal_input = "请输入 URL。";
var image_help = "插入图像\n\n在文本中插入一幅图像。\n例如：[img]http://www.loveyuki.com/images/logo.gif[/img]";
var image_normal = "请输入图像的 URL。";
var flash_help = "插入 Flash\n\n在文本中插入 flash 动画。\n例如：[swf]http://Blog.fz0132.com/images/banner.swf[/swf]";
var flash_normal = "请选择此 Fash 动画 是否自动播放。是：\"Y\" 否：\"N\"。\n此处也可留空。单一页面最多一个。";
var flash_normal_error = "错误：列表格式只能选择输入 \"Y\" 或 \"N\"。";
var flash_normal_input = "请输入 Fash 动画的 URL。";
var wma_help = "插入音频文件\n\n在文本中插入音频文件。\n例如：[wma]http://Blog.fz0132.com/music/music.mp3[/wma]";
var wma_normal = "请选择此音频是否自动播放。是：\"Y\" 否：\"N\"。\n此处也可留空。单一页面最多一个。";
var wma_normal_error = "错误：列表格式只能选择输入 \"Y\" 或 \"N\"。";
var wma_normal_input = "请输入音频文件的 URL。";
var wmv_help = "插入视频文件\n\n在文本中插入视频文件。\n例如：[wmv]http://Blog.fz0132.com/mtv/mtv.wmv[/wmv]";
var wmv_normal = "请选择此视频是否自动播放。是：\"Y\" 否：\"N\"。\n此处也可留空。单一页面最多一个。";
var wmv_normal_error = "错误：列表格式只能选择输入 \"Y\" 或 \"N\"。";
var wmv_normal_input = "请输入视频文件的 URL。";
var ra_help = "插入 RA音频文件\n\n在文本中插入音频文件。\n例如：[ra]http://Blog.fz0132.com/mtv/mtv.ra[/ra]";
var ra_normal = "请选择此 RA音频是否自动播放。是：\"Y\" 否：\"N\"。\n此处也可留空。单一页面最多一个。";
var ra_normal_error = "错误：列表格式只能选择输入 \"Y\" 或 \"N\"。";
var ra_normal_input = "请输入 RA音频文件的 URL。";
var rm_help = "插入RM视频文件\n\n在文本中插入视频文件。\n例如：[rm]http://Blog.fz0132.com/mtv/mtv.rm[/rm]";
var rm_normal = "请选择此 RM视频是否自动播放。是：\"Y\" 否：\"N\"。\n此处也可留空。单一页面最多一个。";
var rm_normal_error = "错误：列表格式只能选择输入 \"Y\" 或 \"N\"。";
var rm_normal_input = "请输入 RM视频文件的 URL。";
var code_help = "插入代码\n\n插入程序或脚本原始代码。\n例如：[code]Response.Write(这里是 Loveyuki's BLOG)[/code]";
var code_normal = "'请输入要插入的代码。";
var list_help = "插入列表\n\n插入可由浏览器显示来的规则列表项。\n例如：\n[list]\n[*]；列表项 #1\n[*]；列表项 #2\n[*]；列表项 #3\n[/list]";
var list_normal = "'请选择列表格式：字母式列表输入 \"A\"；数字式列表输入 \"1\"。此处也可留空。";
var list_normal_error = "错误：列表格式只能选择输入 \"A\" 或 \"1\"。";
var list_normal_input = "请输入列表项目内容，如果留空表示项目结束。";
var sub_help = "插入下标文字\n\n给标签所包围的文本加上下标\n例如: [sub]MY BLOG[/sub]";
var sub_normal = "请输入要加的下标文字。";
var sup_help = "插入上标文字\n\n给标签所包围的文本加上上标\n例如: [sub]MY BLOG[/sub]";
var sup_normal = "请输入要加的上标文字。";
var underline_help = "插入下划线\n\n'给标签所包围的文本加上下划线。\n例如：[u]Loveyuki's BLOG[/u]";
var underline_normal = "请输入要加下划线的文字。";
var light_help = "插入闪光文字\n\n给标签所包围的文本加上闪光\n例如: [light]MY BLOG[/light]";
var light_normal = "请输入要闪光的文字。";
var glow_help = "插入发光文字\n\n给标签所包围的文本加上发光\n例如: [glow]MY BLOG[/glow]";
var glow_normal = "请输入要发光的文字。";
var light_help = "插入闪光文字\n\n给标签所包围的文本加上闪光\n例如: [light]MY BLOG[/light]";
var light_normal = "请输入要闪光的文字。";
var shadow_help = "插入阴影文字\n\n给标签所包围的文本加上阴影\n例如: [glow]MY BLOG[/glow]";
var shadow_normal = "请输入要阴影的文字。";
var mem_help = "插入权限文字\n\n给标签所包围的文本加上权限\n例如: [mem]MY BLOG[/mem]";
var mem_normal = "请输入要权限查看的文字。";
var download_help = "插入权限链接\n\n插入一个权限连接。\n例如：\n[download]http://www.loveyuki.com[/download]\n[download=http://www.loveyuki.com]Loveyuki's BLOG[/download]";
var download_normal = "请输入链接显示的文字。";
var download_normal_input = "请输入 URL。";