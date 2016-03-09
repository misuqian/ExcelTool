# ExcelTool
ExcelTool是一个能根据参考行列快速地进行拼接，分割Excel文件的工具。界面简洁，操作简单
<br>a tool to conditional slice or paste .xls,.xlsx in a quik,simplify way

##条件拼接 Conditional Paste
以paste01.xlsm和paste02.xlsm第一列（A column）进行文件拼接
<br>
![paset01](https://github.com/misuqian/ExcelTool/blob/master/png/paste01.png) **+** 
![paset02](https://github.com/misuqian/ExcelTool/blob/master/png/paste02.png) __=__
![paset_result](https://github.com/misuqian/ExcelTool/blob/master/png/paste_result.png)
<br>
__你只需要根据界面简单设置__ [paste_config](https://github.com/misuqian/ExcelTool/blob/master/png/paste_config.png)

##条件切割 Conditional Slice
以slice_refer.xlsx为源参考，切割slice_body.xlsx
<br>
![slice_body](https://github.com/misuqian/ExcelTool/blob/master/png/slice_body.png) __+__
![slice_refer](https://github.com/misuqian/ExcelTool/blob/master/png/slice_refer.png) __=__
![slice_result01](https://github.com/misuqian/ExcelTool/blob/master/png/slice_result01.png)
<br>__or =__
![slice_result02](https://github.com/misuqian/ExcelTool/blob/master/png/slice_result02.png)
<br>
__你只需要根据界面简单设置__
[slice_config](https://github.com/misuqian/ExcelTool/blob/master/png/slice_config01.png)
<br>
__需要切割到不同文件在这里选择即可__
![slice_config2](https://github.com/misuqian/ExcelTool/blob/master/png/slice_config02.png)

##高级拼接 Advanced Paste
进行文件拼接时，忽略某些行（或者列）不进行拼接
<br>
![paset01](https://github.com/misuqian/ExcelTool/blob/master/png/paste_advance01.png) **+** 
![paset02](https://github.com/misuqian/ExcelTool/blob/master/png/paste_advance02.png) __=__
![paset_result](https://github.com/misuqian/ExcelTool/blob/master/png/paste_advance_result.png)
<br>
__在对应文件导入窗口设置忽略行数（或列数即可）__
![paste_config](https://github.com/misuqian/ExcelTool/blob/master/png/paste_advance_config02.png)
<br>

##高级分割 Advanced Slice
将某些行自动保存到所有的分割文件中,并且以自己为源文件进行切割
<br>
![slice_body](https://github.com/misuqian/ExcelTool/blob/master/png/paste_advance_result.png) __+__
![slice_refer](https://github.com/misuqian/ExcelTool/blob/master/png/paste_advance_result.png) __=__
![slice_result01](https://github.com/misuqian/ExcelTool/blob/master/png/slice_advance_result.png)
<br>
__通过设置忽略行以及参考自己本身即可__
![slice_config](https://github.com/misuqian/ExcelTool/blob/master/png/slice_advance_config.png)

<br>
