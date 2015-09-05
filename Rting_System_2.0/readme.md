文件结构：<br/>
    1.main.py 个人赛 team.py 组队赛<br/>
    2.fileIO.py 与excel发生数据上的交互（使用xlrd和xlwt的数据）以及HTML的生成<br/>
      excel源数据需要：<br/>
        第一个sheet标注为student第一列标注nickname，第二列标注真实姓名<br/>
        建议在第二行标注名为team的sheet，同样的第一列标注nickname，第二列标注真实姓名，之后列跟随队员名，最大人数无限制，但是需要标注为空<br/>
        建议在第三个标注为total的sheet，格式类似于第一个sheet，第一列标注nickname，第二列标注真实姓名，后面至少跟随一组数据，程序按照最后的分数表作为当前默认的rating<br/>
        其他表格直接使用virtual judge导出格式导出到excel即可<br/>
    3.calc.py 通用的计算模块<br/>
    4.board.css作为默认css<br/>
