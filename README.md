# 文件夹构成：<br>
* 1.CmCat为所有的代码所在文件<br>
* 2.CmCat_Export为导出文件所在位置<br>
* 3.CmCat_MySQL为要导入数据库的文件所在位置<br>
* 4.CmCat_Query为示例批量查询文件所在位置<br>
# CmCat文件夹构成：<br>
* 1.core_code为核心代码，包括控件功能与槽函数等<br>
* 2.core_res.qrc为资源文件，包括桌面应用的左上角图标<br>
* 3.CmCat_Icon为控件图标文件所在位置<br>
* 4.CmCat_Res为资源文件所在位置<br>
* 5.CmCat_Ui为所有的UI文件所在位置<br>
* 6.dist中含有打包之后的可执行文件
# 注意：
* 由于git限制大文件无法上传，可执行文件需要自行生成<br>
* 语句：pyinstaller -F -w -i CmCat.ico core_code.py
