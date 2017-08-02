# extraction_contract
本程序适用于Linux系统下抽取doc或docx文件里面的内容<br>
包括 合同名称、合同编号、合同金额、乙方、承办部门、签订时间、承办人、合同类型、拟稿时间、合同份数 等信息<br>
具体介绍如下<br>
函数extract_title 提取 合同名称 <br>
函数extract_number 提取 合同编号 <br>
函数extract_money 提取 合同金额 <br>
函数extract_secondparty 提取 乙方 <br>
函数extract_undertakingdepartment 提取 承办部门 <br>
函数extract_dateofsigning 提取 签订时间 <br>
函数extract_undertaker 提取 承办人 <br>
函数extract_typeofcontract 提取 合同类型 <br>
函数extract_dateprepared 提取 拟稿时间 <br>
函数extract_numberofcontractcopies 提取 合同份数 <br>
主函数main(a,b) 参数a要提取的文件所在的文件夹，提取的信息保存在excel文件中，参数b是excel保存的位置<br>

程序使用示例<br>
调用main函数<br>
main('/home/mylinux/workspace/150823','/home/mylinux/workspace/b.xlsx') <br>
要提取的word文档在Linux 的'/home/mylinux/workspace/150823'文件夹下，提取的信息保存在excel文件中<br>
并且excel 文件命名为b.xlsx,保存在Linux 的'/home/mylinux/workspace/'文件下

