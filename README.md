本代码编写的初衷是语雀文档的迁移，目前已经有作者编写了语雀文档的批量下载，但是下载格式是markdown文件，下载的图片都在同级目录。

由此就产生了需求，想要上传到其他平台并能够像word一样，支持一键解析为云文档。

本程序会将md文件中的pdf链接自动下载到同级目录，由于word不支持插入文档，即使插入了文档，也不能解析到云端，所以直接下载到同级目录。

感谢作者编写了语雀文档批量下载：https://github.com/Be1k0/yuque_document_download/tree/main

配合下载后的md转换为word可以方便解析到云文档（例如飞书）

即可实现类似于语雀转飞书需求。


yuque_document_download经过本人修改，只需要在该程序目录下修改host.txt即可控制登录的空间，例如语雀登录的空间地址：https://www.yuque.com ，host中填写https://immc.yuque.com 即可
