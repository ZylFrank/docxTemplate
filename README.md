## 根据模板生成word文档

对docxtpl的使用

[文档地址](https://docxtpl.readthedocs.io/en/latest/)

其中包含以下几种情况：

- 添加图片
- 根据编码生成条形码图片
- 渲染表格
- 控制表格/段落是否显示
- 使用图片替换占位图片

### 安装环境

```
cd ./docxTemplate
python3 -m venv venv
. venv/bin/activate
pip install -r requirements.txt
python app.py
```

### 脚本运行环境：

- python3
- python-barcode
- Pillow
- docxtpl

## 参数实例

`source_docx_path`: 模版文件路径

`target_docx_path`: 文件输出路径

`barcode_images`: 需要根据条码号生成图片的条码号集合

`images`: 文档中的图片

`content`: 直接以key value形式渲染的字段内容

```
{
    'source_docx_path': './templates/report.docx',
    'target_docx_path': './output/report.docx',
    'barcode_images': [
        {
            'tag': 'barcode_image_1',
            'code': 'BJYJ201900150221',
            'width': 40,
            'height': 20,
        },
    ],
    'images': [
        {
            'tag': 'image1',
            'path': './templates/python.png',
            'width': 100,
            'height': 100,
        },
    ],
    'content': {
        'barcode1': 'BJYJ201900150221',
        'table1': [
            {'index': '1', 'barcode': 'BGG1', 'origin_barcode': 'BGG01', 'name': '北1', 'description': '第1个样品'},
            {'index': '2', 'barcode': 'BGG2', 'origin_barcode': 'BGG02', 'name': '北2', 'description': '第2个样品'},
            {'index': '3', 'barcode': 'BGG3', 'origin_barcode': 'BGG03', 'name': '北3', 'description': '第3个样品'},
            {'index': '4', 'barcode': 'BGG4', 'origin_barcode': 'BGG04', 'name': '北4', 'description': '第4个样品'},
            {'index': '5', 'barcode': 'BGG5', 'origin_barcode': 'BGG05', 'name': '北5', 'description': '第5个样品'},
            {'index': '6', 'barcode': 'BGG6', 'origin_barcode': 'BGG06', 'name': '北6', 'description': '第6个样品'},
        ],
        'table2': [
            {'index': '1', 'barcode': 'BCC1', 'origin_barcode': 'BCC01', 'name': '京1', 'description': '第1个样品'},
            {'index': '2', 'barcode': 'BCC2', 'origin_barcode': 'BCC02', 'name': '京2', 'description': '第2个样品'},
            {'index': '3', 'barcode': 'BCC3', 'origin_barcode': 'BCC03', 'name': '京3', 'description': '第3个样品'},
            {'index': '4', 'barcode': 'BCC4', 'origin_barcode': 'BCC04', 'name': '京4', 'description': '第4个样品'},
            {'index': '5', 'barcode': 'BCC5', 'origin_barcode': 'BCC05', 'name': '京5', 'description': '第5个样品'},
            {'index': '6', 'barcode': 'BCC6', 'origin_barcode': 'BCC06', 'name': '京6', 'description': '第6个样品'},
        ],
        'table3': [
            {'index': '1', 'dc_code': 'BGG1', 'dc_name': '北1', 'dz_code': 'BCC1', 'dz_name': '京1',
             'compare_num': '40',
             'diff_num': '5', 'conclusion': '差异'},
            {'index': '2', 'dc_code': 'BGG2', 'dc_name': '北2', 'dz_code': 'BCC2', 'dz_name': '京2',
             'compare_num': '40',
             'diff_num': '5', 'conclusion': '差异'},
            {'index': '3', 'dc_code': 'BGG3', 'dc_name': '北3', 'dz_code': 'BCC3', 'dz_name': '京3',
             'compare_num': '40',
             'diff_num': '5', 'conclusion': '差异'},
            {'index': '4', 'dc_code': 'BGG4', 'dc_name': '北4', 'dz_code': 'BCC4', 'dz_name': '京4',
             'compare_num': '40',
             'diff_num': '5', 'conclusion': '差异'},
            {'index': '5', 'dc_code': 'BGG5', 'dc_name': '北5', 'dz_code': 'BCC5', 'dz_name': '京5',
             'compare_num': '40',
             'diff_num': '5', 'conclusion': '差异'},
            {'index': '6', 'dc_code': 'BGG6', 'dc_name': '北6', 'dz_code': 'BCC6', 'dz_name': '京6',
             'compare_num': '40',
             'diff_num': '5', 'conclusion': '差异'},
        ],
        'table1_display': True,
        'table_index_1': '1',
        'table2_display': False,
        'table3_display': True,
        'table_index_3': '2',
        'table1_display_data': [
            {'col1': 'table1_row1_col1', 'col2': 'table1_row1_col2'},
            {'col1': 'table1_row1_col1', 'col2': 'table1_row2_col2'},
        ],
        'table3_display_data': [
            {'col1': 'table3_row1_col1', 'col2': 'table3_row1_col2'},
            {'col1': 'table3_row2_col1', 'col2': 'table3_row2_col2'},
        ],
    }
}
```